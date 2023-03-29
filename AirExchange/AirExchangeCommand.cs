/*
 * Copyright (c) <2023> <Misharev Evgeny>
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
 * 1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
 * 2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer 
 *    in the documentation and/or other materials provided with the distribution.
 * 3. Neither the name of the <organization> nor the names of its contributors may be used to endorse or promote products derived 
 *    from this software without specific prior written permission.
 * 4. Redistributions are not allowed to be sold, in whole or in part, for any compensation of any kind.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, 
 * BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. 
 * IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, 
 * OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; 
 * OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT 
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 *
 * Contact: <citrusbim@gmail.com> or <https://web.telegram.org/k/#@MisharevEvgeny>
 */
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AirExchange
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AirExchangeCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<Space> spaceList = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .WhereElementIsNotElementType()
                .Cast<Space>()
                .Where(s => s.Area > 0)
                .ToList();
            if (spaceList.Count == 0)
            {
                TaskDialog.Show("Revit", "В проекте отсутствуют Пространства!");
                return Result.Cancelled;
            }

            //Проверка на наличие параметров
            //Норма воздухообмена
            Parameter airExchangeRateParam = spaceList.First().LookupParameter("Норма воздухообмена");
            if (airExchangeRateParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Норма воздухообмена\"!");
                return Result.Cancelled;
            }

            //По кратности
            Parameter byMultiplicityParam = spaceList.First().LookupParameter("По кратности");
            if (byMultiplicityParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"По кратности\"!");
                return Result.Cancelled;
            }

            //По людям
            Parameter byPeopleParam = spaceList.First().LookupParameter("По людям");
            if (byPeopleParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"По людям\"!");
                return Result.Cancelled;
            }

            //По вредностям
            Parameter byHarmfulnessParam = spaceList.First().LookupParameter("По вредностям");
            if (byHarmfulnessParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"По вредностям\"!");
                return Result.Cancelled;
            }

            //По сан. приборам
            Parameter byPlumbingFixturesParam = spaceList.First().LookupParameter("По сан. приборам");
            if (byPlumbingFixturesParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"По сан. приборам\"!");
                return Result.Cancelled;
            }

            //По балансу
            Parameter byBalanceParam = spaceList.First().LookupParameter("По балансу");
            if (byBalanceParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"По балансу\"!");
                return Result.Cancelled;
            }

            //Кратность_приток
            Parameter multiplicityInflowParam = spaceList.First().LookupParameter("Кратность_приток");
            if (multiplicityInflowParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Кратность_приток\"!");
                return Result.Cancelled;
            }

            //Кратность_приток
            Parameter multiplicityExtractionParam = spaceList.First().LookupParameter("Кратность_вытяжка");
            if (multiplicityExtractionParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Кратность_вытяжка\"!");
                return Result.Cancelled;
            }

            //ADSK_Расчетный приток
            Parameter estimatedInflowParam = spaceList.First().get_Parameter(new Guid("ff939149-328d-421c-93c3-3348a7e55697"));
            if (estimatedInflowParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"ADSK_Расчетный приток\"!");
                return Result.Cancelled;
            }

            //ADSK_Расчетная вытяжка
            Parameter estimatedExtractionParam = spaceList.First().get_Parameter(new Guid("550f0463-71d7-4856-879c-11f9004d5789"));
            if (estimatedExtractionParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"ADSK_Расчетная вытяжка\"!");
                return Result.Cancelled;
            }

            //ADSK_Расчетное количество людей с постоянным пребыванием
            Parameter numberPermanentResidentsParam = spaceList.First().get_Parameter(new Guid("c7feec1c-9972-4277-9bb1-b4ad207f7bca"));
            if (numberPermanentResidentsParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"ADSK_Расчетное количество людей с постоянным пребыванием\"!");
                return Result.Cancelled;
            }

            //Санитарная норма притока на постоянно пребывающего
            Parameter sanitaryNormInflowFoPermanentResidentParam = spaceList.First().LookupParameter("Санитарная норма притока на постоянно пребывающего");
            if (sanitaryNormInflowFoPermanentResidentParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Санитарная норма притока на постоянно пребывающего\"!");
                return Result.Cancelled;
            }

            //Количество писсуаров
            Parameter numberOfUrinalsParam = spaceList.First().LookupParameter("Количество писсуаров");
            if (numberOfUrinalsParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Количество писсуаров\"!");
                return Result.Cancelled;
            }

            //Вытяжка на 1 писсуар
            Parameter extractionOneUrinalParam = spaceList.First().LookupParameter("Вытяжка на 1 писсуар");
            if (extractionOneUrinalParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Вытяжка на 1 писсуар\"!");
                return Result.Cancelled;
            }

            //Количество унитазов
            Parameter numberOfToiletsParam = spaceList.First().LookupParameter("Количество унитазов");
            if (numberOfToiletsParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Количество унитазов\"!");
                return Result.Cancelled;
            }

            //Вытяжка на 1 унитаз
            Parameter extractionOneToiletParam = spaceList.First().LookupParameter("Вытяжка на 1 унитаз");
            if (extractionOneToiletParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Вытяжка на 1 унитаз\"!");
                return Result.Cancelled;
            }

            //Количество душевых кабин 
            Parameter numberOfShowersParam = spaceList.First().LookupParameter("Количество душевых кабин");
            if (numberOfShowersParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Количество душевых кабин\"!");
                return Result.Cancelled;
            }

            //Вытяжка на 1 душ
            Parameter extractionOneShowerParam = spaceList.First().LookupParameter("Вытяжка на 1 душ");
            if (extractionOneShowerParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"Вытяжка на 1 душ\"!");
                return Result.Cancelled;
            }

            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
            Parameter airInflowInfoParam = spaceList.First().LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток");
            if (airInflowInfoParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"ЦИТ_ОВ_Пояснение_воздухообмена_Приток\"!");
                return Result.Cancelled;
            }

            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
            Parameter airExtractionInfoParam = spaceList.First().LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка");
            if (airExtractionInfoParam == null)
            {
                TaskDialog.Show("Revit", "У Пространств отсутствует параметр \"ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка\"!");
                return Result.Cancelled;
            }

            foreach (Space space in spaceList)
            {
                if (doc.GetElement(space.LookupParameter("Норма воздухообмена").AsElementId()) != null)
                {
                    int chek = space.LookupParameter("По кратности").AsInteger()
                        + space.LookupParameter("По вредностям").AsInteger()
                        + space.LookupParameter("По людям").AsInteger()
                        + space.LookupParameter("По сан. приборам").AsInteger()
                        + space.LookupParameter("По балансу").AsInteger();
                    if (chek > 1)
                    {
                        TaskDialog.Show("Revit", $"В Пространстве \"{space.Number} - {space.get_Parameter(BuiltInParameter.ROOM_NAME).AsString()}\" задано больше одного варианта расчета воздухообмена!");
                        return Result.Cancelled;
                    }
                }
            }

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Воздухообмен");
                foreach (Space space in spaceList)
                {
                    if (doc.GetElement(space.LookupParameter("Норма воздухообмена").AsElementId()) != null)
                    {
                        //Объем пространства
#if R2019 || R2020 || R2021
                        double spaceVolume = UnitUtils.ConvertFromInternalUnits(space.Volume, DisplayUnitType.DUT_CUBIC_METERS);
#else
                        double spaceVolume = UnitUtils.ConvertFromInternalUnits(space.Volume, UnitTypeId.CubicMeters);
#endif

                        if (space.LookupParameter("По кратности").AsInteger() == 1)
                        {
#if R2019 || R2020 || R2021
                            //Приток
                            double multiplicityInflow = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Кратность_приток").AsDouble(), DisplayUnitType.DUT_LITERS_PER_SECOND_CUBIC_METER);
                            double estimatedInflow = UnitUtils.ConvertToInternalUnits(RoundFive(spaceVolume * multiplicityInflow), DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR);
                            space.get_Parameter(new Guid("ff939149-328d-421c-93c3-3348a7e55697")).Set(estimatedInflow);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток").Set($"{(int)Math.Round(multiplicityInflow, 0)}");

                            //Вытяжка
                            double multiplicityExtraction = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Кратность_вытяжка").AsDouble(), DisplayUnitType.DUT_LITERS_PER_SECOND_CUBIC_METER);
                            double estimatedExtraction = UnitUtils.ConvertToInternalUnits(RoundFive(spaceVolume * multiplicityExtraction), DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR);
                            space.get_Parameter(new Guid("550f0463-71d7-4856-879c-11f9004d5789")).Set(estimatedExtraction);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{(int)Math.Round(multiplicityExtraction, 0)}");
#else
                            //Приток
                            double multiplicityInflow = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Кратность_приток").AsDouble(), UnitTypeId.CubicMetersPerHourCubicMeter);
                            double estimatedInflow = UnitUtils.ConvertToInternalUnits(RoundFive(spaceVolume * multiplicityInflow), UnitTypeId.CubicMetersPerHour);
                            space.get_Parameter(new Guid("ff939149-328d-421c-93c3-3348a7e55697")).Set(estimatedInflow);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток").Set($"{(int)Math.Round(multiplicityInflow, 0)}");

                            //Вытяжка
                            double multiplicityExtraction = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Кратность_вытяжка").AsDouble(), UnitTypeId.CubicMetersPerHourCubicMeter);
                            double estimatedExtraction = UnitUtils.ConvertToInternalUnits(RoundFive(spaceVolume * multiplicityExtraction), UnitTypeId.CubicMetersPerHour);
                            space.get_Parameter(new Guid("550f0463-71d7-4856-879c-11f9004d5789")).Set(estimatedExtraction);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{(int)Math.Round(multiplicityExtraction, 0)}");
#endif
                        }
                        else if (space.LookupParameter("По людям").AsInteger() == 1)
                        {
#if R2019 || R2020 || R2021
                            //Приток
                            double numberPermanentResidents = space.get_Parameter(new Guid("c7feec1c-9972-4277-9bb1-b4ad207f7bca")).AsInteger();
                            double sanitaryNormInflowFoPermanentResident = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Санитарная норма притока на постоянно пребывающего").AsDouble(), DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR);
                            double estimatedInflow = UnitUtils.ConvertToInternalUnits(RoundFive(numberPermanentResidents * sanitaryNormInflowFoPermanentResident), DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR);
                            space.get_Parameter(new Guid("ff939149-328d-421c-93c3-3348a7e55697")).Set(estimatedInflow);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток").Set($"{sanitaryNormInflowFoPermanentResident} м³/ч на 1 человека");

                            //Вытяжка
                            space.get_Parameter(new Guid("550f0463-71d7-4856-879c-11f9004d5789")).Set(estimatedInflow);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{sanitaryNormInflowFoPermanentResident} м³/ч на 1 человека");

#else
                            //Приток
                            double numberPermanentResidents = space.get_Parameter(new Guid("c7feec1c-9972-4277-9bb1-b4ad207f7bca")).AsInteger();
                            double sanitaryNormInflowFoPermanentResident = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Санитарная норма притока на постоянно пребывающего").AsDouble(), UnitTypeId.CubicMetersPerHour);
                            double estimatedInflow = UnitUtils.ConvertToInternalUnits(RoundFive(numberPermanentResidents * sanitaryNormInflowFoPermanentResident), UnitTypeId.CubicMetersPerHour);
                            space.get_Parameter(new Guid("ff939149-328d-421c-93c3-3348a7e55697")).Set(estimatedInflow);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток").Set($"{sanitaryNormInflowFoPermanentResident} м³/ч на 1 человека");

                            //Вытяжка
                            space.get_Parameter(new Guid("550f0463-71d7-4856-879c-11f9004d5789")).Set(estimatedInflow);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{sanitaryNormInflowFoPermanentResident} м³/ч на 1 человека");
#endif
                        }
                        else if (space.LookupParameter("По сан. приборам").AsInteger() == 1)
                        {
#if R2019 || R2020 || R2021
                            //Приток
                            space.get_Parameter(new Guid("ff939149-328d-421c-93c3-3348a7e55697")).Set(0);
                            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток").Set("");

                            //Вытяжка
                            int numberOfUrinals = space.LookupParameter("Количество писсуаров").AsInteger();
                            double extractionOneUrinal = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Вытяжка на 1 писсуар").AsDouble(), DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR);

                            int numberOfToilets = space.LookupParameter("Количество унитазов").AsInteger();
                            double extractionOneToilet = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Вытяжка на 1 унитаз").AsDouble(), DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR);

                            int numberOfShowers = space.LookupParameter("Количество душевых кабин").AsInteger();
                            double extractionOneShower = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Вытяжка на 1 душ").AsDouble(), DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR);

                            double estimatedExtraction = UnitUtils.ConvertToInternalUnits(RoundFive(numberOfUrinals * extractionOneUrinal
                                + numberOfToilets * extractionOneToilet
                                + numberOfShowers * extractionOneShower), DisplayUnitType.DUT_CUBIC_METERS_PER_HOUR);
                            space.get_Parameter(new Guid("550f0463-71d7-4856-879c-11f9004d5789")).Set(estimatedExtraction);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
                            if (numberOfUrinals > 0 && numberOfToilets > 0 && numberOfShowers > 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneUrinal} м³/ч на 1 писсуар, {extractionOneToilet} м³/ч на 1 унитаз, {extractionOneShower} м³/ч на 1 душевую кабину");
                            }
                            else if (numberOfUrinals > 0 && numberOfToilets > 0 && numberOfShowers == 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneUrinal} м³/ч на 1 писсуар, {extractionOneToilet} м³/ч на 1 унитаз");
                            }
                            else if (numberOfUrinals > 0 && numberOfToilets == 0 && numberOfShowers > 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneUrinal} м³/ч на 1 писсуар, {extractionOneShower} м³/ч на 1 душевую кабину");
                            }
                            else if (numberOfUrinals == 0 && numberOfToilets > 0 && numberOfShowers > 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneToilet} м³/ч на 1 унитаз, {extractionOneShower} м³/ч на 1 душевую кабину");
                            }
                            else if (numberOfUrinals > 0 && numberOfToilets == 0 && numberOfShowers == 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneUrinal} м³/ч на 1 писсуар");
                            }
                            else if (numberOfUrinals == 0 && numberOfToilets > 0 && numberOfShowers == 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneToilet} м³/ч на 1 унитаз");
                            }
                            else if (numberOfUrinals == 0 && numberOfToilets == 0 && numberOfShowers > 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneShower} м³/ч на 1 душевую кабину");
                            }

                            else
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set("");
                            }
#else
                            //Приток
                            space.get_Parameter(new Guid("ff939149-328d-421c-93c3-3348a7e55697")).Set(0);
                            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток").Set("");

                            //Вытяжка
                            int numberOfUrinals = space.LookupParameter("Количество писсуаров").AsInteger();
                            double extractionOneUrinal = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Вытяжка на 1 писсуар").AsDouble(), UnitTypeId.CubicMetersPerHour);

                            int numberOfToilets = space.LookupParameter("Количество унитазов").AsInteger();
                            double extractionOneToilet = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Вытяжка на 1 унитаз").AsDouble(), UnitTypeId.CubicMetersPerHour);

                            int numberOfShowers = space.LookupParameter("Количество душевых кабин").AsInteger();
                            double extractionOneShower = UnitUtils.ConvertFromInternalUnits(space.LookupParameter("Вытяжка на 1 душ").AsDouble(), UnitTypeId.CubicMetersPerHour);

                            double estimatedExtraction = UnitUtils.ConvertToInternalUnits(RoundFive(numberOfUrinals * extractionOneUrinal
                                + numberOfToilets * extractionOneToilet 
                                + numberOfShowers * extractionOneShower), UnitTypeId.CubicMetersPerHour);
                            space.get_Parameter(new Guid("550f0463-71d7-4856-879c-11f9004d5789")).Set(estimatedExtraction);

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
                            if (numberOfUrinals > 0 && numberOfToilets > 0 && numberOfShowers > 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneUrinal} м³/ч на 1 писсуар, {extractionOneToilet} м³/ч на 1 унитаз, {extractionOneShower} м³/ч на 1 душевую кабину");
                            }
                            else if(numberOfUrinals > 0 && numberOfToilets > 0 && numberOfShowers == 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneUrinal} м³/ч на 1 писсуар, {extractionOneToilet} м³/ч на 1 унитаз");
                            }
                            else if (numberOfUrinals > 0 && numberOfToilets == 0 && numberOfShowers > 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneUrinal} м³/ч на 1 писсуар, {extractionOneShower} м³/ч на 1 душевую кабину");
                            }
                            else if (numberOfUrinals == 0 && numberOfToilets > 0 && numberOfShowers > 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneToilet} м³/ч на 1 унитаз, {extractionOneShower} м³/ч на 1 душевую кабину");
                            }
                            else if (numberOfUrinals > 0 && numberOfToilets == 0 && numberOfShowers == 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneUrinal} м³/ч на 1 писсуар");
                            }
                            else if (numberOfUrinals == 0 && numberOfToilets > 0 && numberOfShowers == 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneToilet} м³/ч на 1 унитаз");
                            }
                            else if (numberOfUrinals == 0 && numberOfToilets == 0 && numberOfShowers > 0)
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set($"{extractionOneShower} м³/ч на 1 душевую кабину");
                            }
                            else
                            {
                                space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set("");
                            }
#endif
                        }
                        else if (space.LookupParameter("По вредностям").AsInteger() == 1)
                        {
                            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток").Set("по расчету");

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set("по расчету");
                        }
                        else if (space.LookupParameter("По балансу").AsInteger() == 1)
                        {
                            //ЦИТ_ОВ_Пояснение_воздухообмена_Приток 
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Приток").Set("по балансу");

                            //ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка
                            space.LookupParameter("ЦИТ_ОВ_Пояснение_воздухообмена_Вытяжка").Set("по балансу");
                        }
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("Revit", "Обработка завершена!");
            return Result.Succeeded;
        }
        static double RoundFive(double d)
        {
            if (d % 5 > 2.5)
                d = (int)(d / 5) * 5 + 5;
            else
                d = (int)(d / 5) * 5;

            return d;
        }
    }
}
