using System.Runtime.InteropServices;
using ExcelDna.Integration.CustomUI;
using DerivativeSecuritiesAddIn.Utility;

namespace DerivativeSecuritiesAddIn.Alpha
{
    [ComVisible(true)]
    public class DemoRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string ribbonId) {
            return @"
                  <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' loadImage='LoadImage'>
                  <ribbon>
                    <tabs>
                      <tab id='tab1' label='My Tab'>
                        <group id='group1' label='My Group'>
                          <button id='button1' image='M' size='large' label='My Button' onAction='OnButtonPressed'/>
                        </group >
                      </tab>
                    </tabs>
                  </ribbon>
                </customUI>";
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            TemplateSubBase.TestTemp();
        }
    }
}
