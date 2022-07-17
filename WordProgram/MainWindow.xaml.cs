using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.VisualBasic;
namespace WordProgram
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }


        public static class ExportRomaji
        {
            
            /// <summary>
            /// 指定した Unicode 文字が、０ から ９ までの数字かどうかを示します。
            /// </summary>
            /// <param name="c">評価する Unicode 文字。</param>
            /// <returns>c が数字である場合は true。それ以外の場合は false。</returns>
            public static bool IsFullwidthDigit(char c)
            {
                return '０' <= c && c <= '９';
            }

            /// <summary>
            /// 指定した Unicode 文字が、英字の大文字かどうかを示します。
            /// </summary>
            /// <param name="c">評価する Unicode 文字。</param>
            /// <returns>c が英字の大文字である場合は true。
            /// それ以外の場合は false。</returns>
            public static bool IsUpperLatin(char c)
            {
                //半角英字と全角英字の大文字の時はTrue
                return ('A' <= c && c <= 'Z') || ('Ａ' <= c && c <= 'Ｚ');
            }

            /// <summary>
            /// 指定した Unicode 文字が、英字の小文字かどうかを示します。
            /// </summary>
            /// <param name="c">評価する Unicode 文字。</param>
            /// <returns>c が英字の小文字である場合は true。
            /// それ以外の場合は false。</returns>
            public static bool IsLowerLatin(char c)
            {
                //半角英字と全角英字の小文字の時はTrue
                return ('a' <= c && c <= 'z') || ('ａ' <= c && c <= 'ｚ');
            }


            public static string Parse(string content)
            {
                var result = TextUtility.ToHiragana(content);
                result = HiraganaToAlphabet(result);
                return result;
            }

            public static string HiraganaToAlphabet(string s1)
            {
                string s2 = "";
                for (int i = 0; i < s1.Length; i++)
                {
                    // 小さい文字が含まれる場合
                    if (i + 1 < s1.Length)
                    {
                        // 「っ」が含まれる場合
                        if (s1.Substring(i, 1).CompareTo("っ") == 0)
                        {
                            s2 += HiraganaToAlphabet1(s1.Substring(i + 1, 1)).Substring(0, 1);
                            continue;
                        }


                        // それ以外の小さい文字
                        string s3 = HiraganaToAlphabet1(s1.Substring(i, 2));
                        if (s3.CompareTo("*") != 0)
                        {
                            string s4 = s1.Substring(i, 2);
                            if((System.Text.RegularExpressions.Regex.IsMatch(s4,
                            @"^\p{N}+$")) || (System.Text.RegularExpressions.Regex.IsMatch(s4,
                            @"^[a-zA-Z]+$")) || (System.Text.RegularExpressions.Regex.IsMatch(s4,
                            @"^[ａ-ｚＡ-Ｚ]+$")))
                            {
                                s2 += Strings.StrConv(s4,VbStrConv.Narrow);
                            }
                            else
                            {
                                s2 += s3;
                            }
                            i++;
                            continue;
                        }
                    }
                    string s = s1.Substring(i, 1);

                    if ((System.Text.RegularExpressions.Regex.IsMatch(s,
                    @"^\p{N}+$")) || (System.Text.RegularExpressions.Regex.IsMatch(s,
                    @"^[a-zA-Z]+$")) || (System.Text.RegularExpressions.Regex.IsMatch(s,
                    @"^[ａ-ｚＡ-Ｚ]+$")))
                    {
                        s2 += Strings.StrConv(s, VbStrConv.Narrow,0);
                    }
                    else
                    {
                        s2 += HiraganaToAlphabet1(s1.Substring(i, 1));
                    }
                    
                }
                return s2;
            }
            static string HiraganaToAlphabet1(string s1)
            {
                switch (s1)
                {
                    case "あ": return "a";
                    case "い": return "i";
                    case "う": return "u";
                    case "え": return "e";
                    case "お": return "o";
                    case "か": return "ka";
                    case "き": return "ki";
                    case "く": return "ku";
                    case "け": return "ke";
                    case "こ": return "ko";
                    case "さ": return "sa";
                    case "し": return "shi";
                    case "す": return "su";
                    case "せ": return "se";
                    case "そ": return "so";
                    case "た": return "ta";
                    case "ち": return "chi";
                    case "つ": return "tsu";
                    case "て": return "te";
                    case "と": return "to";
                    case "な": return "na";
                    case "に": return "ni";
                    case "ぬ": return "nu";
                    case "ね": return "ne";
                    case "の": return "no";
                    case "は": return "ha";
                    case "ひ": return "hi";
                    case "ふ": return "hu";
                    case "へ": return "he";
                    case "ほ": return "ho";
                    case "ま": return "ma";
                    case "み": return "mi";
                    case "む": return "mu";
                    case "め": return "me";
                    case "も": return "mo";
                    case "や": return "ya";
                    case "ゆ": return "yu";
                    case "よ": return "yo";
                    case "ら": return "ra";
                    case "り": return "ri";
                    case "る": return "ru";
                    case "れ": return "re";
                    case "ろ": return "ro";
                    case "わ": return "wa";
                    case "を": return "wo";
                    case "ん": return "n";
                    case "が": return "ga";
                    case "ぎ": return "gi";
                    case "ぐ": return "gu";
                    case "げ": return "ge";
                    case "ご": return "go";
                    case "ざ": return "za";
                    case "じ": return "ji";
                    case "ず": return "zu";
                    case "ぜ": return "ze";
                    case "ぞ": return "zo";
                    case "だ": return "da";
                    case "ぢ": return "ji";
                    case "づ": return "du";
                    case "で": return "de";
                    case "ど": return "do";
                    case "ば": return "ba";
                    case "び": return "bi";
                    case "ぶ": return "bu";
                    case "べ": return "be";
                    case "ぼ": return "bo";
                    case "ぱ": return "pa";
                    case "ぴ": return "pi";
                    case "ぷ": return "pu";
                    case "ぺ": return "pe";
                    case "ぽ": return "po";
                    case "きゃ": return "kya";
                    case "きぃ": return "kyi";
                    case "きゅ": return "kyu";
                    case "きぇ": return "kye";
                    case "きょ": return "kyo";
                    case "しゃ": return "sha";
                    case "しぃ": return "syi";
                    case "しゅ": return "shu";
                    case "しぇ": return "she";
                    case "しょ": return "sho";
                    case "ちゃ": return "cha";
                    case "ちぃ": return "cyi";
                    case "ちゅ": return "chu";
                    case "ちぇ": return "che";
                    case "ちょ": return "cho";
                    case "にゃ": return "nya";
                    case "にぃ": return "nyi";
                    case "にゅ": return "nyu";
                    case "にぇ": return "nye";
                    case "にょ": return "nyo";
                    case "ひゃ": return "hya";
                    case "ひぃ": return "hyi";
                    case "ひゅ": return "hyu";
                    case "ひぇ": return "hye";
                    case "ひょ": return "hyo";
                    case "みゃ": return "mya";
                    case "みぃ": return "myi";
                    case "みゅ": return "myu";
                    case "みぇ": return "mye";
                    case "みょ": return "myo";
                    case "りゃ": return "rya";
                    case "りぃ": return "ryi";
                    case "りゅ": return "ryu";
                    case "りぇ": return "rye";
                    case "りょ": return "ryo";
                    case "ぎゃ": return "gya";
                    case "ぎぃ": return "gyi";
                    case "ぎゅ": return "gyu";
                    case "ぎぇ": return "gye";
                    case "ぎょ": return "gyo";
                    case "じゃ": return "ja";
                    case "じぃ": return "ji";
                    case "じゅ": return "ju";
                    case "じぇ": return "je";
                    case "じょ": return "jo";
                    case "ぢゃ": return "dya";
                    case "ぢぃ": return "dyi";
                    case "ぢゅ": return "dyu";
                    case "ぢぇ": return "dye";
                    case "ぢょ": return "dyo";
                    case "びゃ": return "bya";
                    case "びぃ": return "byi";
                    case "びゅ": return "byu";
                    case "びぇ": return "bye";
                    case "びょ": return "byo";
                    case "ぴゃ": return "pya";
                    case "ぴぃ": return "pyi";
                    case "ぴゅ": return "pyu";
                    case "ぴぇ": return "pye";
                    case "ぴょ": return "pyo";
                    case "ぐぁ": return "gwa";
                    case "ぐぃ": return "gwi";
                    case "ぐぅ": return "gwu";
                    case "ぐぇ": return "gwe";
                    case "ぐぉ": return "gwo";
                    case "つぁ": return "tsa";
                    case "つぃ": return "tsi";
                    case "つぇ": return "tse";
                    case "つぉ": return "tso";
                    case "ふぁ": return "fa";
                    case "ふぃ": return "fi";
                    case "ふぇ": return "fe";
                    case "ふぉ": return "fo";
                    case "うぁ": return "wha";
                    case "うぃ": return "whi";
                    case "うぅ": return "whu";
                    case "うぇ": return "whe";
                    case "うぉ": return "who";
                    case "ヴぁ": return "va";
                    case "ヴぃ": return "vi";
                    case "ヴ": return "vu";
                    case "ヴぇ": return "ve";
                    case "ヴぉ": return "vo";
                    case "でゃ": return "dha";
                    case "でぃ": return "dhi";
                    case "でゅ": return "dhu";
                    case "でぇ": return "dhe";
                    case "でょ": return "dho";
                    case "てゃ": return "tha";
                    case "てぃ": return "thi";
                    case "てゅ": return "thu";
                    case "てぇ": return "the";
                    case "てょ": return "tho";
                    default: return "*";
                }
            }
        }

        public static class TextUtility
        {
            // IFELanguage2 Interface ID
            //[Guid("21164102-C24A-11d1-851A-00C04FCC6B14")]
            [ComImport]
            [Guid("019F7152-E6DB-11d0-83C3-00C04FDDB82E")]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IFELanguage
            {
                int Open();
                int Close();
                int GetJMorphResult(uint dwRequest, uint dwCMode, int cwchInput, [MarshalAs(UnmanagedType.LPWStr)] string pwchInput, IntPtr pfCInfo, out object ppResult);
                int GetConversionModeCaps(ref uint pdwCaps);
                int GetPhonetic([MarshalAs(UnmanagedType.BStr)] string @string, int start, int length, [MarshalAs(UnmanagedType.BStr)] out string result);
                int GetConversion([MarshalAs(UnmanagedType.BStr)] string @string, int start, int length, [MarshalAs(UnmanagedType.BStr)] out string result);
            }

            public static string ToHiragana(string source)
            {
                IFELanguage? ifelang = null;
                try
                {
                    var type = Type.GetTypeFromProgID("MSIME.Japan") ?? throw new Exception();
                    ifelang = Activator.CreateInstance(type) as IFELanguage ?? throw new Exception();
                    int hr = ifelang.Open();
                    if (hr != 0)
                    {
                        throw Marshal.GetExceptionForHR(hr) ?? throw new Exception();
                    }
                    string yomigana;
                    hr = ifelang.GetPhonetic(source, 1, -1, out yomigana);
                    if (hr != 0)
                    {
                        throw Marshal.GetExceptionForHR(hr) ?? throw new Exception();
                    }
                    return yomigana;
                }
                finally
                {
                    ifelang?.Close();
                }
            }
        }
        public class TextConverter : IDisposable
        {
            // IFELanguage2 Interface ID
            //[Guid("21164102-C24A-11d1-851A-00C04FCC6B14")]
            [ComImport]
            [Guid("019F7152-E6DB-11d0-83C3-00C04FDDB82E")]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            public interface IFELanguage
            {
                int Open();
                int Close();
                int GetJMorphResult(uint dwRequest, uint dwCMode, int cwchInput, [MarshalAs(UnmanagedType.LPWStr)] string pwchInput, IntPtr pfCInfo, out object ppResult);
                int GetConversionModeCaps(ref uint pdwCaps);
                int GetPhonetic([MarshalAs(UnmanagedType.BStr)] string @string, int start, int length, [MarshalAs(UnmanagedType.BStr)] out string result);
                int GetConversion([MarshalAs(UnmanagedType.BStr)] string @string, int start, int length, [MarshalAs(UnmanagedType.BStr)] out string result);
            }

            private IFELanguage _ifelang;

            public TextConverter()
            {
                var type = Type.GetTypeFromProgID("MSIME.Japan") ?? throw new Exception();
                _ifelang = Activator.CreateInstance(type) as IFELanguage ?? throw new Exception();
                int hr = _ifelang.Open();
                if (hr != 0)
                {
                    throw Marshal.GetExceptionForHR(hr) ?? throw new Exception($"{hr} is not error");
                }
            }

            public void Dispose()
            {
                _ifelang?.Close();
            }

            public string ToHiragana(string source)
            {
                int hr = _ifelang.GetPhonetic(source, 1, -1, out string yomigana);
                if (hr != 0)
                {
                    throw Marshal.GetExceptionForHR(hr) ?? throw new Exception($"{hr} is not error");
                }
                return yomigana;
            }
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            Result1.Text = "";
            Result2.Text = "";
            Result3.Text = "";
            var sentenses = origin.Text.Split('、');
            Result1.Text = "";
            foreach (var sentenses2 in sentenses)
            {
                var sentenses3 = sentenses2.Split('。');
                foreach (string sentenses4 in sentenses3)
                {
                    var sentenses5 = sentenses4.Split(' ');
                    foreach(string sentenses6 in sentenses5)
                    {
                        var result = TextUtility.ToHiragana(sentenses6);
                        Result1.Text += result;
                        
                    }
                    
                }
            }
            var finallyresult = ExportRomaji.HiraganaToAlphabet(Result1.Text).Replace("*", "");
            Result2.Text = finallyresult;

            string[] a1 = finallyresult.ToCharArray().Select(c => new string(c, 1)).ToArray();
            var Keys = new List<string>();
            var KeyWord = new List<List<string>>();
            var Value = new List<List<int>>();

            for (int i = 0; i < a1.Length - 1; i++)
            {
                if (!Keys.Contains(a1[i]))
                {
                    Keys.Add(a1[i]);
                    Value.Add(new List<int>());
                    KeyWord.Add(new List<string>());
                }

                for (int l = 0; l < Keys.Count; l++)
                {
                    if (Keys[l] == a1[i])
                    {
                        if (!KeyWord[l].Contains(a1[i + 1]))
                        {
                            KeyWord[l].Add(a1[i + 1]);
                            Value[l].Add(0);
                        }

                        for (int k = 0; k < KeyWord[l].Count; k++)
                        {
                            if (KeyWord[l][k] == a1[i + 1])
                            {
                                Value[l][k] += 1;
                            }
                        }

                        break;
                    }

                }

            }
            Result3.Text = "";
            for (int i = 0; i < Keys.Count - 1; i++)
            {
                Result3.Text += $"{Keys[i]}の後に続く単語の確率\n";
                for (int j = 0; j < Value[i].Count; j++)
                {
                    int all = 0;
                    foreach (var value in Value[i])
                    {
                        all += value;
                    }
                    Result3.Text += $"{KeyWord[i][j]}キー {(100 * (decimal)Value[i][j] / all).ToString()}%\n";
                    await Task.Delay(20);

                }
            }



        }
    }
}
