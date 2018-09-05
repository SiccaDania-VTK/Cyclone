Imports System.IO
Imports System.Math
'Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading

'------- Korrel groepen in de inlaat stroom------
Public Structure Korrel_struct
    Public dia As Double            'Particle diameter [mu]
    Public aandeel As Double        'Aandeel in de inlaat stroom [% weight]
    Public verlies As Double        'verlies (nietgevangen) [-]
End Structure

Public Class Form1
    Public korrel(22) As Korrel_struct    '22 korrel groepen

    'Type AC;Inlaatbreedte;Inlaathoogte;Inlaatlengte;Inlaat hartmaat;Inlaat afschuining;
    'Uitlaat keeldia inw.;Uitlaat flensdiameter inw.;Lengte insteekpijp inw.;
    'Lengte romp + conus;Lengte romp;Lengte conus;Dia_conus / 3P-pijp;Lengte 3P-pijp;Lengte 3P-conus;Kleine dia 3P-conus",

    Public Shared cyl_dimensions() As String = {
    "AC-300;0.34;0.77;0.6;0.63;0.3;0.68;0.68;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-350;0.32;0.7;0.6;0.617;0.3;0.63;0.63;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-435;0.282;0.64;0.6;0.6;0.3;0.56;0.56;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-550;0.25;0.57;0.6;0.58;0.3;0.45;0.56;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-750;0.216;0.486;0.6;0.57;0.3;0.365;0.56;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-850;0.203;0.457;0.6;0.564;0.3;0.307;0.428;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-1850;0.136;0.31;0.6;0.53;0.3;0.15;0.25;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25"}

    'FRACTIONELE VERLIESCIJFERS volgens norm 503
    'Verlies aan de hand van de deeltjes grootte
    'min dia[mm];max dia[mm],AC300;AC350;AC435;AC550;AC750;AC850;AC850;AC1850;AC1850
    Public Shared frac_verlies() As String = {
     "0;2;97.00;95.00;87.00;80.00;75.00;70.00;60.00;60.00;30.00",
    "2;4;76.00;70.00;60.00;47.00;40.00;30.00;26.00;20.00;7.00",
    "4;6;54.00;48.00;40.00;30.00;25.00;16.00;9.00;12.00;3.00",
    "6;8;45.00;32.00;21.00;17.00;14.00;8.70;3.70;6.50;1.50",
    "8;10;36.00;22.00;12.00;10.00;8.00;5.15;1.18;4.00;1.00",
    "10;12;29.00;16.00;8.00;6.50;4.60;3.40;1.10;2.50;0.60",
    "12;14;20.50;11.00;5.50;3.50;2.70;2.40;0.65;1.70;0.45",
    "14;16;14.00;7.50;3.00;2.20;1.60;1.60;0.50;1.20;0.35",
    "16;18;11.00;5.50;2.20;1.40;1.10;1.10;0.35;0.85;0.25",
    "18;20;8.40;4.00;1.60;0.90;0.75;0.75;0.25;0.60;0.20",
    "20;25;5.50;2.50;1.00;0.45;0.40;0.40;0.16;0.40;0.15",
    "25;30;4.20;1.60;0.50;0.18;0.15;0.15;0.10;0.15;0.08",
    "30;100;3.20;0.80;0.15;0.07;0.06;0.05;0.045;0.05;0.04"}

    'FRACTIONELE VERLIESCIJFERS volgens norm 503 
    '[mm]; <  2; 2 - 4; 4 - 6; 6 - 8; 8 - 10; 10 - 12; 12 - 14; 14 - 16; 16 - 18; 18 - 20; 20 - 25; 25 - 30; > 30
    Public Shared frac_verlies2() As String = {
    "AC300;97.00;76.00;54.00;45.00;36.00;29.00;20.50;14.00;11.00;8.40;5.50;4.20;3.20",
    "AC350;95.00;70.00;48.00;32.00;22.00;16.00;11.00;7.50;5.50;4.00;2.50;1.60;0.80",
    "AC435;87.00;60.00;40.00;21.00;12.00;8.00;5.50;3.00;2.20;1.60;1.00;0.50;0.15",
    "AC550;80.00;47.00;30.00;17.00;10.00;6.50;3.50;2.20;1.40;0.90;0.45;0.18;0.07",
    "AC750;75.00;40.00;25.00;14.00;8.00;4.60;2.70;1.60;1.10;0.75;0.40;0.15;0.06",
    "AC850;70.00;30.00;16.00;8.70;5.15;3.40;2.40;1.60;1.10;0.75;0.40;0.15;0.05",
    "AC850;60.00;26.00;9.00;3.70;1.18;1.10;0.65;0.50;0.35;0.25;0.16;0.10;0.045",
    "AC1850;60.00;20.00;12.00;6.50;4.00;2.50;1.70;1.20;0.85;0.60;0.40;0.15;0.05",
    "AC1850;30.00;7.00;3.00;1.50;1.00;0.60;0.45;0.35;0.25;0.20;0.15;0.08;0.04"}


    'Nieuwe reken methode, verdeling volgens Weibull
    'm1,k1,a1 als d < d_krit
    'm2,k2,a2 als d > d_krit
    ' type;dkrit;m1;k1;a1;m2;k2;a2
    Public Shared rekenlijnen() As String = {
    "AC300;     12.2;   1.15;   7.457;  1.005;      8.5308;     1.6102; 0.4789",
    "AC350;     10.2;   1.0;    5.3515; 1.0474;     4.4862;     2.4257; 0.6472",
    "AC435;     8.93;   0.69;   4.344;  1.139;      4.2902;     1.3452; 0.5890",
    "AC550;     8.62;   0.527;  3.4708; 0.9163;     3.3211;     1.7857; 0.7104",
    "AC750;     8.3;    0.50;   2.8803; 0.8355;     4.0940;     1.0519; 0.6010",
    "AC850;     7.8;    0.52;   1.9418; 0.73705;    -0.1060;    2.0197; 0.7077",
    "AC850+afz; 10;     0.5187; 1.6412; 0.8386;     4.2781;     0.06777;0.3315",
    "AC1850;    9.3;    0.50;   1.1927; 0.5983;     -0.196;     1.3687; 0.6173",
    "AC1850+afz;10.45;  0.4617; 0.2921; 0.4560;     -0.2396;    0.1269; 0.3633"}

    Public weerstand_coef(7) As Double               'Poly Coefficients, Polynomial regression


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        'De weerstandscoefficient volgt uit het cycloon type
        weerstand_coef = {7, 7, 7, 7, 7.5, 9.5, 14.5}

        For hh = 0 To (cyl_dimensions.Length - 1)  'Fill combobox9 Insulation data
            words = cyl_dimensions(hh).Split(";")
            ComboBox1.Items.Add(words(0))
        Next hh

        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 2                 'Select Cyclone type
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles button1.Click, TabPage1.Enter, numericUpDown5.ValueChanged, numericUpDown9.ValueChanged, numericUpDown8.ValueChanged, numericUpDown7.ValueChanged, numericUpDown6.ValueChanged, numericUpDown12.ValueChanged, numericUpDown11.ValueChanged, numericUpDown10.ValueChanged, ComboBox1.SelectedValueChanged, numericUpDown3.ValueChanged, numericUpDown2.ValueChanged, numericUpDown14.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown15.ValueChanged, CheckBox1.CheckedChanged
        Dim words() As String
        Dim cyl_dim(20), db As Double
        Dim in_hoog, in_breed, Body_dia, Flow, inlet_velos, delta_p, K_waarde As Double
        Dim ro_gas, ro_particle, visco, wc As Double

        Dim Totaal_korrel_verlies As Double  'Berekende verlies

        If (ComboBox1.SelectedIndex > -1) Then      'Prevent exceptions
            words = cyl_dimensions(ComboBox1.SelectedIndex).Split(";")
            For hh = 1 To 15
                cyl_dim(hh) = words(hh)
            Next

            db = numericUpDown5.Value           'Body diameter
            in_hoog = cyl_dim(1) * db           '[m]
            in_breed = cyl_dim(2) * db          '[m]
            Body_dia = numericUpDown5.Value     '[m]
            Flow = NumericUpDown1.Value         '[m3/s]
            ro_gas = numericUpDown3.Value       '[kg/m3]
            ro_particle = numericUpDown2.Value  '[kg/m3]
            visco = numericUpDown14.Value       '[cPoise]

            '----------- inlaat snelheid ---------------------
            inlet_velos = Flow / (in_breed * in_hoog)

            '----------- Pressure loss cyclone----------------------
            wc = weerstand_coef(ComboBox1.SelectedIndex)
            delta_p = 0.5 * ro_gas * inlet_velos ^ 2 * wc

            '----------- K_waarde-----------------------------------
            K_waarde = db * 2000 * visco * 16 / (ro_particle * 0.0181 * inlet_velos)
            K_waarde = Sqrt(K_waarde)


            TextBox22.Text = Round(Calc_verlies(NumericUpDown15.Value) * 100, 1).ToString       'verlies getal[%]

            '----------- presenteren ----------------------------------
            TextBox1.Text = Round(in_hoog, 2).ToString              'inlaat breedte
            TextBox2.Text = Round(in_breed, 2).ToString             'Inlaat hoogte
            TextBox3.Text = Round(cyl_dim(3) * db, 2).ToString      'Inlaat lengte
            TextBox4.Text = Round(cyl_dim(4) * db, 2).ToString      'Inlaat hartmaat
            TextBox5.Text = Round(cyl_dim(5) * db, 2).ToString      'Inlaat afschuining

            TextBox6.Text = Round(cyl_dim(6) * db, 2).ToString      'Uitlaat keeldia inw.
            TextBox7.Text = Round(cyl_dim(7) * db, 2).ToString      'Uitlaat flensdiameter inw.

            TextBox8.Text = Round(cyl_dim(8) * db, 2).ToString      'Lengte insteekpijp inw.

            TextBox9.Text = Round(cyl_dim(9) * db, 2).ToString      'Lengte romp + conus
            TextBox10.Text = Round(cyl_dim(10) * db, 2).ToString    'Lengte romp
            TextBox11.Text = Round(cyl_dim(11) * db, 2).ToString    'Lengte çonus

            TextBox12.Text = Round(cyl_dim(12) * db, 2).ToString    'Dia_conus / 3P-pijp
            TextBox13.Text = Round(cyl_dim(13) * db, 2).ToString    'Lengte 3P-pijp

            TextBox14.Text = Round(cyl_dim(14) * db, 2).ToString    'Dia_conus / 3P-pijp
            TextBox15.Text = Round(cyl_dim(15) * db, 2).ToString    'Lengte 3P-pijp

            TextBox16.Text = Round(inlet_velos, 1).ToString         'inlaat snelheid
            TextBox17.Text = Round(delta_p, 0).ToString             'Pressure loss

            TextBox23.Text = Round(K_waarde, 4).ToString           'Stokes waarde tov Standaard cycloon
            Draw_chart()
            '---------- Check speed ---------------
            If inlet_velos < 12 Or inlet_velos > 25 Then
                TextBox16.BackColor = Color.Red
            Else
                TextBox16.BackColor = Color.LightGreen
            End If

            '--------- Inlet korrel data -----------
            korrel(0).dia = 1
            korrel(1).dia = 3
            korrel(2).dia = 5
            korrel(3).dia = 8
            korrel(4).dia = 14
            korrel(5).dia = 24
            korrel(6).dia = 40
            korrel(7).dia = 75

            korrel(0).aandeel = numericUpDown6.Value / 100  'Percentale van de inlaat stof belasting
            korrel(1).aandeel = numericUpDown7.Value / 100
            korrel(2).aandeel = numericUpDown8.Value / 100
            korrel(3).aandeel = numericUpDown9.Value / 100
            korrel(4).aandeel = numericUpDown10.Value / 100
            korrel(5).aandeel = numericUpDown11.Value / 100
            korrel(6).aandeel = numericUpDown12.Value / 100

            '---- moet opgeteld 100% zijn --------------
            korrel(7).aandeel = 1
            For h = 0 To 6
                korrel(7).aandeel -= korrel(h).aandeel
            Next
            numericUpDown13.Value = korrel(7).aandeel * 100

            If korrel(7).aandeel < 0 Or korrel(7).aandeel > 1 Then
                numericUpDown13.BackColor = Color.Red
            Else
                numericUpDown13.BackColor = Color.LightGreen
            End If

            '--------- overall resultaat --------------------
            Totaal_korrel_verlies = 0
            For h = 0 To 7
                korrel(h).verlies = Calc_verlies(korrel(h).dia)
                Totaal_korrel_verlies += korrel(0).aandeel * korrel(h).verlies
            Next h

            TextBox24.Text = Round(korrel(0).verlies * 100, 1).ToString
            TextBox25.Text = Round(korrel(1).verlies * 100, 1).ToString
            TextBox26.Text = Round(korrel(2).verlies * 100, 1).ToString
            TextBox27.Text = Round(korrel(3).verlies * 100, 1).ToString
            TextBox28.Text = Round(korrel(4).verlies * 100, 1).ToString
            TextBox29.Text = Round(korrel(5).verlies * 100, 1).ToString
            TextBox30.Text = Round(korrel(6).verlies * 100, 1).ToString
            TextBox31.Text = Round(korrel(7).verlies * 100, 1).ToString
            TextBox32.Text = Round(Totaal_korrel_verlies * 100, 1).ToString 'Totaal verlies [-]
            TextBox33.Text = Round(Totaal_korrel_verlies * NumericUpDown4.Value, 2).ToString 'Totaal verlies [kg/s]
        End If
    End Sub
    '-------- Bereken het verlies getal -----------
    '----- de input is de korrel grootte-----------
    Private Function Calc_verlies(korrel_g As Double)
        Dim words() As String
        Dim dia_krit, fac_m, fac_a, fac_k, verlies, kwaarde As Double


        Double.TryParse(TextBox23.Text, kwaarde)

        '-------------- korrelgrootte factoren ------
        words = rekenlijnen(ComboBox1.SelectedIndex).Split(";")

        dia_krit = words(1)

        '-------- de grafieken zijn in 2 delen gesplits voor hogere nouwkeurigheid----------
        If korrel_g < dia_krit Then
            fac_m = words(2)
            fac_k = words(3)
            fac_a = words(4)
        Else
            fac_m = words(5)
            fac_k = words(6)
            fac_a = words(7)
        End If

        verlies = -((korrel_g / kwaarde - fac_m) / fac_k) ^ fac_a
        verlies = Math.E ^ verlies

        '---------- present------------------
        TextBox18.Text = Round(dia_krit, 3).ToString            'diameter_kritisch
        TextBox19.Text = Round(fac_m, 3).ToString               'faktor-m
        TextBox20.Text = Round(fac_k, 3).ToString               'faktor-kappa
        TextBox21.Text = Round(fac_a, 3).ToString               'faktor-a
        Return (verlies)
    End Function


    Private Sub Draw_chart()
        '-------
        Dim s_points(100, 2) As Double
        Dim h As Integer

        Chart1.Series.Clear()
        Chart1.ChartAreas.Clear()
        Chart1.Titles.Clear()
        Chart1.ChartAreas.Add("ChartArea0")

        ' For h = 0 To 1
        Chart1.Series.Add("Series" & h.ToString)
        Chart1.Series(h).ChartArea = "ChartArea0"
        Chart1.Series(h).ChartType = DataVisualization.Charting.SeriesChartType.Line
        '  Chart1.Series(schets_no).Name = (Tschets(schets_no).Tname)
        Chart1.Series(h).BorderWidth = 1
        Chart1.Series(h).IsVisibleInLegend = False
        ' Next

        Chart1.Titles.Add("Verlies Curve")
        Chart1.ChartAreas("ChartArea0").AxisX.Title = "particle dia [mu]"

        Chart1.ChartAreas("ChartArea0").AxisY.Title = "Loss [%] (niet gevangen)"
        Chart1.ChartAreas("ChartArea0").AxisY.Minimum = 0       'Loss
        Chart1.ChartAreas("ChartArea0").AxisY.Maximum = 100     'Loss
        Chart1.ChartAreas("ChartArea0").AxisY.Interval = 10     'Interval

        If CheckBox1.Checked Then
            Chart1.ChartAreas("ChartArea0").AxisX.IsLogarithmic = True
            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 1     'Particle size
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = 100   'Particle size
        Else
            Chart1.ChartAreas("ChartArea0").AxisX.IsLogarithmic = False
            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0     'Particle size
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = 40    'Particle size
        End If

        '----- now calc --------------------------
        For h = 0 To 100
            s_points(h, 0) = h                                   'Particle diameter [mu]
            s_points(h, 1) = calc_verlies(s_points(h, 0)) * 100  'Loss [%]
        Next

        '------ now present-------------
        For h = 0 To 40 - 1   'Fill line chart
            Chart1.Series(0).Points.AddXY(s_points(h, 0), s_points(h, 1))
        Next h
    End Sub
End Class
