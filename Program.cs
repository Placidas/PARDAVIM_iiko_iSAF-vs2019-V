using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Linq;
using System.Diagnostics;
using System.Windows;
using System.Xml.Linq;

namespace PARDAVIM_iiko_iSAF
{
    class Program
    {
        public static string IM_KODAS = "304823023";

        public static List<String> NUMERIS = new List<String>();
        public static List<String> DATA = new List<String>();
        public static List<String> KLIENTAS = new List<String>();
        public static List<String> KLIENTAS_UNIKALUS = new List<String>();
        public static List<String> PVM = new List<String>();
        public static List<String> KODAS = new List<String>();
        public static List<decimal> SUMA = new List<decimal>();
        public static List<decimal> PVM_SUMA = new List<decimal>();
        public static List<decimal> PVM_TARIFAS = new List<decimal>();
        public static string Numeris, Data, Klientas, Pvm, Kodas, START_DATE, END_DATE;
        public static decimal Suma = 0, Pvm_suma = 0, Pvm_tarifas = 0, Suma_viso = 0, Pvm_viso = 0;
        public static int Field_Numeris = -1,
                          Field_Data = -1,
                          Field_Klientas = -1,
                          Field_Im_kodas = -1,
                          Field_PVM_kodas = -1,
                          Field_PVM_suma = -1,
                          Field_PVM_tarifas = -1,
                          Field_Suma = -1;
        public static decimal GET_DECIMAL(string txt)
        {
            decimal dec = 0, i = 0; char c = '?';
            for (; txt.Length > 0 && char.IsDigit(c = txt[0]); txt = txt.Substring(1)) if (char.IsDigit(c = txt[0])) dec = dec * 10 + c - '0';
            if (txt.Length > 0)
            {
                if (c == ',' || c == '.')
                {
                    txt = txt.Substring(1);
                    if (txt.Length > 0)
                    {
                        if (char.IsDigit(c = txt[0]))
                        {
                            i = c - '0'; dec += i / 10;
                            if (txt.Length > 0)
                            {
                                txt = txt.Remove(0, 1);
                                if (txt.Length > 0) if (char.IsDigit(c = txt[0])) i = c - '0'; dec += i / 100;
                            }
                        }
                    }
                }
            }
            return dec;
        }

        static void Main()
        {
            int NR, NR_KLIENTU, nr, i, len;
            string FILE_NAME, Line = "", tmp = "";
            START_DATE = "9"; END_DATE = "0";
            try
            {
                string directory = Directory.GetCurrentDirectory();
                string[] files = Directory.GetFiles(directory, "Ataskaita pagal sąskaitas faktūras *.txt");
                if (files.Count() == 0)
                {
                    MessageBox.Show("Kataloge\r\n\r\n" + directory + "\r\n\r\nnėra ataskaitų pagal sąskaitas faktūras");
                    return;
                }
                FILE_NAME = files[0];
                len = FILE_NAME.Length;
                foreach (string file in files)
                {
                    if (file.Length < len) continue;
                    if (String.Compare(FILE_NAME.Substring(len - 17, 4), file.Substring(len - 17, 4)) > 0) continue;                        // Metai
                    if (String.Compare(FILE_NAME.Substring(len - 17, 4), file.Substring(len - 17, 4)) < 0) { FILE_NAME = file; continue; }
                    if (String.Compare(FILE_NAME.Substring(len - 20, 2), file.Substring(len - 20, 2)) > 0) continue;                        // Mėnuo
                    if (String.Compare(FILE_NAME.Substring(len - 20, 2), file.Substring(len - 20, 2)) < 0) { FILE_NAME = file; continue; }
                    if (String.Compare(FILE_NAME.Substring(len - 23, 2), file.Substring(len - 23, 2)) > 0) continue;                        // Diena
                    if (String.Compare(FILE_NAME.Substring(len - 23, 2), file.Substring(len - 23, 2)) < 0) { FILE_NAME = file; continue; }
                    if (String.Compare(FILE_NAME.Substring(len - 12, 2), file.Substring(len - 12, 2)) > 0) continue;                        // Valandos
                    if (String.Compare(FILE_NAME.Substring(len - 12, 2), file.Substring(len - 12, 2)) < 0) { FILE_NAME = file; continue; }
                    if (String.Compare(FILE_NAME.Substring(len - 9, 2), file.Substring(len - 9, 2)) > 0) continue;                          // Minutės
                    FILE_NAME = file;
                }
                tmp = FILE_NAME.Substring(len - 17, 4) + "-" + FILE_NAME.Substring(len - 20, 2) + "-" + FILE_NAME.Substring(len - 23, 2) + " " +
                      FILE_NAME.Substring(len - 12, 2) + ":" + FILE_NAME.Substring(len - 9, 2);
                if (MessageBox.Show("Surastas failas su paskutine data yra: " + tmp + "\r\n\r\n                      Ar formuoti duomenis ?\r\n\r\n", "",
                    MessageBoxButtons.YesNo).ToString() == "No") return;

                StreamReader rd = new System.IO.StreamReader(FILE_NAME, System.Text.Encoding.GetEncoding("Windows-1257"));
                for (len = 0; len < 3; len++)
                    if ((Line = rd.ReadLine()) == null)
                    {
                        MessageBox.Show("Klaidos antraštėje faile\r\n" + FILE_NAME);
                        return;
                    }
                for (nr = 0; nr < 20; nr++)
                {
                    if ((len = Line.IndexOf("\t")) < 0) break;
                    tmp = Line.Remove(len);
                    if (tmp.IndexOf("Numeris") >= 0) Field_Numeris = nr;
                    else if (tmp.IndexOf("Data") >= 0) Field_Data = nr;
                    else if (tmp.IndexOf("Klientas") >= 0) Field_Klientas = nr;
                    else if (tmp.IndexOf("m. kodas") >= 0) Field_Im_kodas = nr;
                    else if (tmp.IndexOf("PVM kodas") >= 0) Field_PVM_kodas = nr;
                    else if (tmp.IndexOf("PVM suma") >= 0) Field_PVM_suma = nr;
                    else if (tmp.IndexOf("PVM tarifas") >= 0) Field_PVM_tarifas = nr;
                    else if (tmp.IndexOf("Suma") >= 0) Field_Suma = nr;
                    Line = Line.Substring(len + 1);
                }
                if (Field_Numeris < 0)      { MessageBox.Show("Faile '" + FILE_NAME + "' nėra lauko 'Numeris'"); return; }
                if (Field_Data < 0)         { MessageBox.Show("Faile '" + FILE_NAME + "' nėra lauko 'Data'"); return; }
                if (Field_Klientas < 0)     { MessageBox.Show("Faile '" + FILE_NAME + "' nėra lauko 'Klientas'"); return; }
                if (Field_Im_kodas < 0)     { MessageBox.Show("Faile '" + FILE_NAME + "' nėra lauko 'Įm. kodas'"); return; }
                if (Field_PVM_kodas < 0)    { MessageBox.Show("Faile '" + FILE_NAME + "' nėra lauko 'PVM kodas'"); return; }
                if (Field_PVM_suma < 0)     { MessageBox.Show("Faile '" + FILE_NAME + "' nėra lauko 'PVM suma EUR'"); return; }
                if (Field_PVM_tarifas < 0)  { MessageBox.Show("Faile '" + FILE_NAME + "' nėra lauko 'PVM tarifas'"); return; }
                if (Field_Suma < 0)         { MessageBox.Show("Faile '" + FILE_NAME + "' nėra lauko 'Suma'"); return; }

                for (NR = NR_KLIENTU = 0; ;)
                {
                    if ((Line = rd.ReadLine()) == null) break;
                    Numeris = Data = Klientas = Kodas = Pvm = "";
                    Pvm_suma = Pvm_tarifas = Suma = -1;

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 0) Numeris = tmp;
                    else if (Field_Data == 0) Data = tmp;
                    else if (Field_Klientas == 0) Klientas = tmp;
                    else if (Field_Im_kodas == 0) Kodas = tmp;
                    else if (Field_PVM_kodas == 0) Pvm = tmp;
                    else if (Field_PVM_suma == 0) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 0) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 0) Suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    Line = Line.Substring(len + 1); if ((len = Line.IndexOf("\t")) < 0) continue; tmp = Line.Remove(len);

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 1) Numeris = tmp;
                    else if (Field_Data == 1) Data = tmp;
                    else if (Field_Klientas == 1) Klientas = tmp;
                    else if (Field_Im_kodas == 1) Kodas = tmp;
                    else if (Field_PVM_kodas == 1) Pvm = tmp;
                    else if (Field_PVM_suma == 1) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 1) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 1) Suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    Line = Line.Substring(len + 1); if ((len = Line.IndexOf("\t")) < 0) continue; tmp = Line.Remove(len);

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 2) Numeris = tmp;
                    else if (Field_Data == 2) Data = tmp;
                    else if (Field_Klientas == 2) Klientas = tmp;
                    else if (Field_Im_kodas == 2) Kodas = tmp;
                    else if (Field_PVM_kodas == 2) Pvm = tmp;
                    else if (Field_PVM_suma == 2) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 2) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 2) Suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    Line = Line.Substring(len + 1); if ((len = Line.IndexOf("\t")) < 0) continue; tmp = Line.Remove(len);

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 3) Numeris = tmp;
                    else if (Field_Data == 3) Data = tmp;
                    else if (Field_Klientas == 3) Klientas = tmp;
                    else if (Field_Im_kodas == 3) Kodas = tmp;
                    else if (Field_PVM_kodas == 3) Pvm = tmp;
                    else if (Field_PVM_suma == 3) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 3) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 3) Suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    Line = Line.Substring(len + 1); if ((len = Line.IndexOf("\t")) < 0) continue; tmp = Line.Remove(len);

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 4) Numeris = tmp;
                    else if (Field_Data == 4) Data = tmp;
                    else if (Field_Klientas == 4) Klientas = tmp;
                    else if (Field_Im_kodas == 4) Kodas = tmp;
                    else if (Field_PVM_kodas == 4) Pvm = tmp;
                    else if (Field_PVM_suma == 4) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 4) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 4) Suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    Line = Line.Substring(len + 1); if ((len = Line.IndexOf("\t")) < 0) continue; tmp = Line.Remove(len);

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 5) Numeris = tmp;
                    else if (Field_Data == 5) Data = tmp;
                    else if (Field_Klientas == 5) Klientas = tmp;
                    else if (Field_Im_kodas == 5) Kodas = tmp;
                    else if (Field_PVM_kodas == 5) Pvm = tmp;
                    else if (Field_PVM_suma == 5) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 5) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 5) Suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    Line = Line.Substring(len + 1); if ((len = Line.IndexOf("\t")) < 0) continue; tmp = Line.Remove(len);

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 6) Numeris = tmp;
                    else if (Field_Data == 6) Data = tmp;
                    else if (Field_Klientas == 6) Klientas = tmp;
                    else if (Field_Im_kodas == 6) Kodas = tmp;
                    else if (Field_PVM_kodas == 6) Pvm = tmp;
                    else if (Field_PVM_suma == 6) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 6) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 6) Suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    Line = Line.Substring(len + 1); if ((len = Line.IndexOf("\t")) < 0) continue; tmp = Line.Remove(len);

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 7) Numeris = tmp;
                    else if (Field_Data == 7) Data = tmp;
                    else if (Field_Klientas == 7) Klientas = tmp;
                    else if (Field_Im_kodas == 7) Kodas = tmp;
                    else if (Field_PVM_kodas == 7) Pvm = tmp;
                    else if (Field_PVM_suma == 7) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 7) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 7) Suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    Line = Line.Substring(len + 1); if ((len = Line.IndexOf("\t")) < 0) continue; tmp = Line.Remove(len);

                    if ((len = Line.IndexOf("\t")) < 0) continue;
                    tmp = Line.Remove(len);
                    if (Field_Numeris == 8) Numeris = tmp;
                    else if (Field_Data == 8) Data = tmp;
                    else if (Field_Klientas == 8) Klientas = tmp;
                    else if (Field_Im_kodas == 8) Kodas = tmp;
                    else if (Field_PVM_kodas == 8) Pvm = tmp;
                    else if (Field_PVM_suma == 8) Pvm_suma = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_PVM_tarifas == 8) Pvm_tarifas = GET_DECIMAL(tmp.Replace(" ", ""));
                    else if (Field_Suma == 8) Suma = GET_DECIMAL(tmp.Replace(" ", ""));

                    if (Numeris == "" || Data == "" || Klientas == "") continue;
                    if (Pvm_tarifas < 0 || Pvm_suma < 0 || Suma < 0) continue;
                    if (Klientas.Remove(1) == "\"" && Klientas.Substring(Klientas.Length - 1) == "\"") Klientas = Klientas.Substring(1, Klientas.Length - 2);
                    if (String.Compare(START_DATE, Data) > 0) START_DATE = Data;
                    if (String.Compare(END_DATE, Data) < 0) END_DATE = Data;

                    NUMERIS.Add(Numeris); DATA.Add(Data); KLIENTAS.Add(Klientas); KLIENTAS_UNIKALUS.Add(Klientas); PVM.Add(Pvm); KODAS.Add(Kodas);
                    SUMA.Add(Suma); PVM_SUMA.Add(Pvm_suma); PVM_TARIFAS.Add(Pvm_tarifas); Suma_viso += Suma; Pvm_viso += Pvm_suma; NR++;
                }
                START_DATE = START_DATE.Remove(8) + "01"; // Visada pirma diena
                string menuo = END_DATE.Substring(5, 2);
                if (menuo == "01" || menuo == "03" || menuo == "05" || menuo == "07" || menuo == "08" || menuo == "10" || menuo == "12") END_DATE = END_DATE.Remove(8) + "31";
                else if (menuo != "02") END_DATE = END_DATE.Remove(8) + "30"; else END_DATE = END_DATE.Remove(8) + "28";

                if (NR == 0)
                {
                    MessageBox.Show("Dokumente nėra duomenų");
                    return;
                }

                StreamWriter wd = new System.IO.StreamWriter(directory + "\\iSAF_klientai.csv", false, System.Text.Encoding.GetEncoding("Windows-1257"));
                wd.WriteLine("Klientas;Kodas;PVM kodas");

                for (nr = 0; nr < NR; nr++)
                {
                    if (KLIENTAS_UNIKALUS[nr] == "") continue;
                    wd.WriteLine(KLIENTAS[nr] + ";" + KODAS[nr] + ";" + PVM[nr]);
                    for (i = nr + 1; i < NR; i++) if (KLIENTAS[nr] == KLIENTAS[i] && KODAS[nr] == KODAS[i] && PVM[nr] == PVM[i]) KLIENTAS_UNIKALUS[i] = "";
                    NR_KLIENTU++;
                }
                wd.Close();

                StreamWriter WD = new System.IO.StreamWriter(directory + "\\iSAF_" + START_DATE + "-" + END_DATE + ".xml", false, System.Text.Encoding.GetEncoding("Windows-1257"));
                WD.WriteLine("<?xml version=\"1.0\" encoding=\"windows-1257\"?>");
                WD.WriteLine("<iSAFFile xmlns=\"http://www.vmi.lt/cms/imas/isaf\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
                WD.WriteLine("  <Header>");
                WD.WriteLine("    <FileDescription>");
                WD.WriteLine("      <FileVersion>iSAF1.2</FileVersion>");
                WD.WriteLine("      <FileDateCreated>" + DateTime.Now.ToString("yyyy-MM-dd") + "T" + DateTime.Now.ToString("HH:mm") + ":00Z</FileDateCreated>");
                WD.WriteLine("      <DataType>P</DataType>");
                WD.WriteLine("      <SoftwareCompanyName>Prestos projektai, UAB</SoftwareCompanyName>");
                WD.WriteLine("      <SoftwareName>PrestProj</SoftwareName>");
                WD.WriteLine("      <SoftwareVersion>1.0</SoftwareVersion>");
                WD.WriteLine("      <RegistrationNumber>" + IM_KODAS + "</RegistrationNumber>");
                WD.WriteLine("      <NumberOfParts>1</NumberOfParts>");
                WD.WriteLine("      <PartNumber>1</PartNumber>");
                WD.WriteLine("      <SelectionCriteria>");
                WD.WriteLine("        <SelectionStartDate>" + START_DATE + "</SelectionStartDate>");
                WD.WriteLine("        <SelectionEndDate>" + END_DATE + "</SelectionEndDate>");
                WD.WriteLine("      </SelectionCriteria>");
                WD.WriteLine("    </FileDescription>");
                WD.WriteLine("  </Header>");
                WD.WriteLine("  <SourceDocuments>");
                WD.WriteLine("    <SalesInvoices>");
                for (nr = 0; nr < NR; nr++)
                {
                    WD.WriteLine("      <Invoice>");
                    WD.WriteLine("        <InvoiceNo>" + NUMERIS[nr] + "</InvoiceNo>");
                    WD.WriteLine("        <CustomerInfo>");
                    if (KODAS[nr] == "") KODAS[nr] = "ND";
                    WD.WriteLine("          <CustomerID>" + KODAS[nr] + "</CustomerID>");
                    if (PVM[nr] == "") PVM[nr] = "ND";
                    WD.WriteLine("          <VATRegistrationNumber>" + PVM[nr] + "</VATRegistrationNumber>");
                    WD.WriteLine("          <RegistrationNumber>" + KODAS[nr] + "</RegistrationNumber>");
                    WD.WriteLine("          <Country>LT</Country>");
                    WD.WriteLine("          <Name>" + KLIENTAS[nr].Replace("&", "&amp;") + "</Name>");
                    WD.WriteLine("        </CustomerInfo>");
                    WD.WriteLine("        <InvoiceDate>" + DATA[nr] + "</InvoiceDate>");
                    WD.WriteLine("        <InvoiceType>SF</InvoiceType>");
                    WD.WriteLine("        <SpecialTaxation/>");
                    WD.WriteLine("        <References/>");
                    WD.WriteLine("        <VATPointDate>" + DATA[nr] + "</VATPointDate>");
                    WD.WriteLine("        <DocumentTotals>");
                    for (string CURR_NUMERIS = NUMERIS[nr], CURR_KODAS = KODAS[nr]; nr < NR; nr++)
                    {
                        if (PVM_SUMA[nr] == 0) continue;
                        WD.WriteLine("          <DocumentTotal>");
                        WD.WriteLine("            <TaxableValue>" + (SUMA[nr] - PVM_SUMA[nr]).ToString("0.00").Replace(",", ".") + "</TaxableValue>");
                        if (PVM_TARIFAS[nr] == 21)
                        {
                            WD.WriteLine("            <TaxCode>PVM1</TaxCode>");
                            WD.WriteLine("            <TaxPercentage>21</TaxPercentage>");
                            WD.WriteLine("            <Amount>" + PVM_SUMA[nr].ToString("0.00").Replace(",", ".") + "</Amount>");
                        }
                        else
                        if (PVM_TARIFAS[nr] == 9)
                        {
                            WD.WriteLine("            <TaxCode>PVM2</TaxCode>");
                            WD.WriteLine("            <TaxPercentage>9</TaxPercentage>");
                            WD.WriteLine("            <Amount>" + PVM_SUMA[nr].ToString("0.00").Replace(",", ".") + "</Amount>");
                        }
                        else
                        {
                            WD.WriteLine("            <TaxCode>PVM5</TaxCode>");
                            WD.WriteLine("            <TaxPercentage>0</TaxPercentage>");
                            WD.WriteLine("            <Amount>0.00</Amount>");
                        }
                        WD.WriteLine("            <VATPointDate2>" + DATA[nr] + "</VATPointDate2>");
                        WD.WriteLine("          </DocumentTotal>");
                        if (nr >= (NR - 1)) break;
                        if (CURR_NUMERIS != NUMERIS[nr + 1] || CURR_KODAS != KODAS[nr + 1]) break;
                    }
                    WD.WriteLine("        </DocumentTotals>");
                    WD.WriteLine("      </Invoice>");
                }
                WD.WriteLine("    </SalesInvoices>");
                WD.WriteLine("  </SourceDocuments>");
                WD.WriteLine("</iSAFFile>");
                WD.Close();

                MessageBox.Show("Suformuota\r\n\r\n  dokumentų: " + NR.ToString() + "\r\n  suma:            " + Suma_viso.ToString("0.00") +
                    "\r\n  PVM suma:   " + Pvm_viso.ToString("0.00") + "\r\n  klientų:          " + NR_KLIENTU.ToString());
            }
            catch (Exception e)
            {
                MessageBox.Show("Sistemos klaida\r\n\r\n" + e.ToString());

            }
        }
    }
}
