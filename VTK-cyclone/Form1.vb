Imports System.IO
Imports System.Math
'Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading

Public Class Form1

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

    'FRACTIONELE VERLIESCIJFERS volgens norm 503  [%]
    Public Shared frac_verlies() As String =
    {"[mm];AC300;AC350;AC435;AC550;AC750;AC850;AC850;AC1850;AC1850",
    "<  2;97.00;95.00;87.00;80.00;75.00;70.00;60.00;60.00;30.00",
    "2 - 4;76.00;70.00;60.00;47.00;40.00;30.00;26.00;20.00;7.00",
    "4 - 6;54.00;48.00;40.00;30.00;25.00;16.00;9.00;12.00;3.00",
    "6 - 8;45.00;32.00;21.00;17.00;14.00;8.70;3.70;6.50;1.50",
    "8 - 10;36.00;22.00;12.00;10.00;8.00;5.15;1.18;4.00;1.00",
    "10 - 12;29.00;16.00;8.00;6.50;4.60;3.40;1.10;2.50;0.60",
    "12 - 14;20.50;11.00;5.50;3.50;2.70;2.40;0.65;1.70;0.45",
    "14 - 16;14.00;7.50;3.00;2.20;1.60;1.60;0.50;1.20;0.35",
    "16 - 18;11.00;5.50;2.20;1.40;1.10;1.10;0.35;0.85;0.25",
    "18 - 20;8.40;4.00;1.60;0.90;0.75;0.75;0.25;0.60;0.20",
    "20 - 25;5.50;2.50;1.00;0.45;0.40;0.40;0.16;0.40;0.15",
    "25 - 30;4.20;1.60;0.50;0.18;0.15;0.15;0.10;0.15;0.08",
    "> 30;3.20;0.80;0.15;0.07;0.06;0.05;0.045;0.05;0.04"}

    'Nieuwe reken methode, verdeling volgens Weibull

    'm1,k1,a1 als d < d_krit
    'm2,k2,a2 als d > d_krit

    Public Shared rekenlijnen() As String = {
    "type;dkrit;m1;k1;a1;m2;k2;a2",
    "AC300;12.2;1.15;7.457;1.005;85.308;16.102;.4789",
    "AC350;10.2;1.0;53.515;10.474;44.862;24.257;0.6472",
    "AC435;8.93;0.69;4.344;1.139;42.902;13.452;0.5890",
    "AC550;8.62;0.527;34.708;0.9163;33.211;17.857;0.7104",
    "AC750;8.3;0.50;28.803;0.8355;40.940;10.519;0.6010",
    "AC850;7.8;0.52;19.418;0.73705;-.1060;20.197;0.7077",
    "AC850+afz;10;0.5187;16.412;0.8386;42.781;0.06777;0.3315",
    "AC1850;9.3;0.50;11.927;0.5983;-0.196;13.687;0.6173",
    "AC1850+afz;10.45;0.4617;0.2921;0.4560;-0.2396;0.1269;0.3633"}



    Public weerstand_coef(7) As Double               'Poly Coefficients, Polynomial regression


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        'De weerstandscoefficient volgt uit het cycloon type
        weerstand_coef = {7, 7, 7, 7, 7.5, 9.5, 14.5}

        For hh = 1 To (cyl_dimensions.Length - 1)  'Fill combobox9 Insulation data
            words = cyl_dimensions(hh).Split(";")
            ComboBox1.Items.Add(words(0))
        Next hh

        If ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 2                 'Select Cyclone type
        End If
    End Sub

    Private Sub button1_Click(sender As Object, e As EventArgs) Handles button1.Click, TabPage1.Enter, numericUpDown5.ValueChanged, numericUpDown9.ValueChanged, numericUpDown8.ValueChanged, numericUpDown7.ValueChanged, numericUpDown6.ValueChanged, numericUpDown13.ValueChanged, numericUpDown12.ValueChanged, numericUpDown11.ValueChanged, numericUpDown10.ValueChanged, ComboBox1.SelectedValueChanged, numericUpDown4.ValueChanged, numericUpDown3.ValueChanged, numericUpDown2.ValueChanged, numericUpDown14.ValueChanged, NumericUpDown1.ValueChanged
        Dim words() As String
        Dim cyl_dim(20), db As Double
        Dim in_hoog, in_breed, Body_dia, Flow, inlet_velos, delta_p As Double
        Dim ro_gas, wc As Double

        If (ComboBox1.SelectedIndex > -1) Then      'Prevent exceptions
            words = cyl_dimensions(ComboBox1.SelectedIndex).Split(";")
            For hh = 1 To 15
                cyl_dim(hh) = words(hh)
            Next

            db = numericUpDown5.Value           'Body diameter
            in_hoog = cyl_dim(1) * db
            in_breed = cyl_dim(2) * db
            Body_dia = numericUpDown5.Value     '[m]
            Flow = NumericUpDown1.Value         '[m3/s]
            ro_gas = numericUpDown3.Value       '[kg/m3]

            '----------- inlaat snelheid ---------------------
            inlet_velos = Flow / (in_breed * in_hoog)

            '----------- Pressure loss cyclone----------------------
            wc = weerstand_coef(ComboBox1.SelectedIndex)
            delta_p = 0.5 * ro_gas * inlet_velos ^ 2 * wc


            '----------- presenteren ----------------------------------
            TextBox1.Text = Round(in_hoog, 2).ToString  'inlaat breedte
            TextBox2.Text = Round(in_breed, 2).ToString  'Inlaat hoogte
            TextBox3.Text = Round(cyl_dim(3) * db, 2).ToString  'Inlaat lengte
            TextBox4.Text = Round(cyl_dim(4) * db, 2).ToString  'Inlaat hartmaat
            TextBox5.Text = Round(cyl_dim(5) * db, 2).ToString  'Inlaat afschuining

            TextBox6.Text = Round(cyl_dim(6) * db, 2).ToString  'Uitlaat keeldia inw.
            TextBox7.Text = Round(cyl_dim(7) * db, 2).ToString  'Uitlaat flensdiameter inw.

            TextBox8.Text = Round(cyl_dim(8) * db, 2).ToString  'Lengte insteekpijp inw.

            TextBox9.Text = Round(cyl_dim(9) * db, 2).ToString  'Lengte romp + conus
            TextBox10.Text = Round(cyl_dim(10) * db, 2).ToString 'Lengte romp
            TextBox11.Text = Round(cyl_dim(11) * db, 2).ToString 'Lengte çonus

            TextBox12.Text = Round(cyl_dim(12) * db, 2).ToString 'Dia_conus / 3P-pijp
            TextBox13.Text = Round(cyl_dim(13) * db, 2).ToString 'Lengte 3P-pijp

            TextBox14.Text = Round(cyl_dim(14) * db, 2).ToString 'Dia_conus / 3P-pijp
            TextBox15.Text = Round(cyl_dim(15) * db, 2).ToString 'Lengte 3P-pijp

            TextBox16.Text = Round(inlet_velos, 2).ToString     'inlaat snelheid
            TextBox17.Text = Round(delta_p, 2).ToString         'Pressure loss
        End If





    End Sub

End Class
