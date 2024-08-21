using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Microsoft.Win32.SafeHandles;


public static class Prompt
{
    public static string ShowDialog(string text, string caption)
    {
        Form prompt = new Form()
        {
            Width = 500,
            Height = 200,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            Text = caption,
            StartPosition = FormStartPosition.CenterScreen,
            BackColor = Color.FromArgb(45, 45, 48),
            ForeColor = Color.White
        };

        Label textLabel = new Label() { Left = 50, Top = 20, Text = text, AutoSize = true };
        TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400, BackColor = Color.Black, ForeColor = Color.White };

        Button confirmation = new Button() { Text = "OK", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK, BackColor = Color.Black, ForeColor = Color.White };
        confirmation.FlatAppearance.BorderColor = Color.Gray;
        confirmation.FlatStyle = FlatStyle.Flat;

        prompt.Controls.Add(textBox);
        prompt.Controls.Add(confirmation);
        prompt.Controls.Add(textLabel);
        prompt.AcceptButton = confirmation;

        return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
    }
}

public class TabData
{
    public string Path { get; set; }
    public int NetworkIndex { get; set; }
    public Rectangle CloseButton { get; set; } // 閉じるボタンの矩形を追加
}

// スクリプトの情報を保持するクラス
public class ScriptInfo
{
    public string FullPath { get; set; }
    public string FileName { get; set; }

    // ComboBoxに表示する文字列を制御します
    public override string ToString()
    {
        return FileName;
    }
}