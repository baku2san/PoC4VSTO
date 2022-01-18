using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PoC4VSTO
{
    public partial class MainRibbon
    {
        private void MainRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
        enum TypeValue
        {
            Formula = 0,	// 数式
            Number = 1,		// 数値
            String = 2,		// 文字列 (テキスト)
            Bool = 4,		// 論理値 (True または False)
            Range = 8,		// セル参照 (Range オブジェクト)
            Errors = 16,	// #N/A などのエラー値
            Array = 64,		// 値の配列
        }
        private void ButtonPoC4InputBox_Click(object sender, RibbonControlEventArgs e)
        {
            var result = Globals.ThisAddIn.Application.InputBox(TypeValue.Formula.ToString(), "Type 比較", Type.Missing, 50, 50, Type.Missing, Type.Missing, TypeValue.Formula);
            MessageBox.Show(result);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var result = Globals.ThisAddIn.Application.InputBox(TypeValue.Number.ToString(), "Type 比較", Type.Missing, 50, 50, Type.Missing, Type.Missing, TypeValue.Number);
            MessageBox.Show(result.ToString());
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            var result = Globals.ThisAddIn.Application.InputBox(TypeValue.String.ToString(), "Type 比較", "default", 50, 50, Type.Missing, Type.Missing, TypeValue.String);
            MessageBox.Show(result);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            var result = Globals.ThisAddIn.Application.InputBox(TypeValue.Bool.ToString(), "Type 比較", Type.Missing, 50, 50, Type.Missing, Type.Missing, TypeValue.Bool);
            MessageBox.Show(result.ToString());
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            var result = Globals.ThisAddIn.Application.InputBox(TypeValue.Range.ToString(), "Type 比較", Type.Missing, 50, 50, Type.Missing, Type.Missing, TypeValue.Range);
            MessageBox.Show(result.Address.ToString());
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            var result = Globals.ThisAddIn.Application.InputBox(TypeValue.Errors.ToString(), "Type 比較", Type.Missing, 50, 50, Type.Missing, Type.Missing, TypeValue.Errors);
            MessageBox.Show(result.ToString());
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {
            var result = Globals.ThisAddIn.Application.InputBox(TypeValue.Array.ToString(), "Type 比較", Type.Missing, 50, 50, Type.Missing, Type.Missing, TypeValue.Array);
            MessageBox.Show(result.GetType().ToString());
        }
    }
}
