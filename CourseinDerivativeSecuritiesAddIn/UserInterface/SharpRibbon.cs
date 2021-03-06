﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using DerivativeSecuritiesAddIn.Alpha;
using DerivativeSecuritiesAddIn.ObjectSystem;
using DerivativeSecuritiesAddIn.Utility;
using ExcelDna.Integration.CustomUI;

namespace DerivativeSecuritiesAddIn.UserInterface
{
    [ComVisible(true)]
    public class SharpRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) {
            return @"
                  <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' loadImage='LoadImage'>
                  <ribbon>
                    <tabs>
                      <tab id='tab1' label='SHARP'>
                        <group id='group1'>
                          <button id='btnAbout' image='About' size='large' label='About Author' onAction='About'/>
                          <button id='btnUpdate' image='Update' size='large' label='Check Update' onAction='CheckUpdate'/>
                        </group>
                        <group id='group2' label='Insert'>
                          <dynamicMenu id='dynamicMenu1' image='Calculator' size='large' label = 'Calculator' getContent = 'GetTempContent' />
                          <button id='btnDemo' size='large' label='+1s' onAction='DemoObj'/>
                        </group>
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        public void About(IRibbonControl control) => Process.Start(
            "https://github.com/gokurakujoudo/");
        public void CheckUpdate(IRibbonControl control) => Process.Start(
            "https://github.com/gokurakujoudo/CourseinDerivativeSecuritiesAddIn/tree/master/CourseinDerivativeSecuritiesAddIn/bin/Debug/Alone");

        public string GetTempContent(IRibbonControl control)
        {
            var sb = new StringBuilder($@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >");
            foreach (var item in TempFuncMenu) 
                sb.Append(AddMenu(item.Key, item.Value, nameof(OnTempPressed)));
            sb.Append(@"</menu>");
            return sb.ToString();
        }

        public string GetTempTitle(IRibbonControl control) { return control.Id; }

        private static readonly Dictionary<string, Action> TempFuncDist = Reflection.GetTempsActions();
        private static readonly Dictionary<string, List<string>> TempFuncMenu=Reflection.GetTempsStrs();

        public void OnTempPressed(IRibbonControl control) => TempFuncDist[control.Id.ToUpper()]?.Invoke();


        private static string AddMenu(string title, IEnumerable<string> items, string action) {
            var sb = new StringBuilder();
            //sb.Append($@"<menu id='{title}' getTitle=""GetTempTitle"" />");
            sb.Append($@"<menuSeparator id='{title}' getTitle=""GetTempTitle"" />");
            foreach (var item in items)
                sb.Append(AddMenuItemButton(item, action));
            //sb.Append(@"</menu>");
            return sb.ToString();
        }

        private static string AddMenuItemButton(string label, string action) => $@"<button id=""{label}"" label=""{label}"" onAction=""{action}"" />";

        public void DemoObj(IRibbonControl control) {
            var temp = new SharpObjectTemplate("DEMO") {
                PropertyList = {{"AGE", 100}},
                PropertyDescription = {{"AGE", "double"}},
                ComputedPropertyList = {{"NEWAGE", so => so["AGE"].ToDouble() + 1}}
            };
            var df = DataFrameHelper.ProvideTemplate();
            SharpObjectExcelTemplate.CreateObjTemp(df);
        }
    }
}
