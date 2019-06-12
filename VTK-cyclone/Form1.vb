Imports System.Globalization
Imports System.IO
Imports System.Management
Imports System.Math
Imports System.Text
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word

'------- Korrel groepen in de inlaat stroom------
Public Structure Korrel_struct
    Public dia_small As Double          'Particle diameter [mu]
    Public dia_big As Double            'Particle diameter [mu]
    Public dia_ave As Double            'Particle diameter [mu]
    Public class_wght_cum_pro As Double 'group_weight_cum in de inlaat stroom [% weight]
    Public class_wght_pro As Double     'group_weight in de inlaat stroom [% weight]
    Public class_wght_kg As Double      'group_weight in de inlaat stroom [kg]
    Public verlies As Double            'verlies (niet gevangen) [-]
End Structure
'Variables used by GvG in calculation
Public Structure GvG_Calc_struct
    Public dia As Double            'Particle diameter [mu]
    Public d_ave As Double          'Average diameter [mu]
    Public d_ave_K As Double        'Average diam/K_stokes [-]
    Public loss_overall As Double   'Overall Corrected
    Public loss_overall_C As Double 'Overall loss Corrected
    Public catch_chart As Double    '[%] for chart
    Public i_grp As Double          'Groepnummer
    Public i_d1 As Double           'Interpolatie dia volgens Rosin Rammler
    Public i_d2 As Double           'Interpolatie dia
    Public i_p1 As Double           'Interpolatie
    Public i_p2 As Double           'Interpolatie
    Public i_k As Double            'Interpolatie
    Public i_m As Double            'Interpolatie
    Public psd_cum As Double        'Partice Size Distribution cummulatief
    Public psd_cump As Double       '[%] PSD cummulatief
    Public psd_dif As Double        '[%] PSD diff
    Public loss_abs As Double       '
    Public loss_abs_C As Double     '
End Structure

Public Class Form1
    Public korrel_grp(22) As Korrel_struct    '22 korrel groepen
    Public guus(150) As GvG_Calc_struct
    Public _K_stokes As Double

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


    'Nieuwe reken methode, verdeling volgens Weibull verdeling
    'm1,k1,a1 als d < d_krit
    'm2,k2,a2 als d > d_krit
    'type; d/krit; m1; k1; a1; m2; k2; a2
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

    Public weerstand_coef_air(7) As Double      'Inlet-air pressure loss calculation
    Public weerstand_coef_dust(7) As Double     'Inlet-air pressure loss calculation

    '----------- directory's-----------
    Dim dirpath_Eng As String = "N:\Engineering\VBasic\Cyclone_sizing_input\"
    Dim dirpath_Rap As String = "N:\Engineering\VBasic\Cyclone_rapport_copy\"
    Dim dirpath_tmp As String = "C:\Tmp\"
    Dim ProcID As Integer = Process.GetCurrentProcess.Id
    Dim dirpath_Temp As String = "C:\Temp\" & ProcID.ToString


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hh, life_time, i As Integer
        Dim words() As String
        Dim separators() As String = {";"}
        Dim Pro_user, HD_number As String
        Dim nu, nu2 As Date
        Dim user_list As New List(Of String)
        Dim hard_disk_list As New List(Of String)
        Dim pass_name As Boolean = False
        Dim pass_disc As Boolean = False

        '------ allowed users with hard disc id's -----
        user_list.Add("user")
        hard_disk_list.Add("058F63646471")      'Privee PC, graslaan25

        user_list.Add("GerritP")
        hard_disk_list.Add("S2R6NX0H740154H")  'VTK PC, GP

        user_list.Add("GerritP")
        hard_disk_list.Add("0008_0D02_003E_0FBB.")       'VTK laptop, GP

        user_list.Add("FredKo")
        hard_disk_list.Add("JR10006P02Y6EE")    'VTK laptop, FKo

        user_list.Add("VittorioS")
        hard_disk_list.Add("002427108605")      'VTK laptop, Vittorio

        user_list.Add("keess")
        hard_disk_list.Add("002410146654")      'VTK laptop, KeesS

        user_list.Add("JanK")
        hard_disk_list.Add("0025_38B4_71B4_88FC.") 'VTK laptop, Jank

        user_list.Add("JeroenA")
        hard_disk_list.Add("171095402070")       'VTK desktop, Jeroen

        user_list.Add("JeroenA")
        hard_disk_list.Add("170228801578")       'VTK laptop, Jeroen disk 1
        hard_disk_list.Add("MCDBM1M4F3QRBEH6")   'VTK laptop, Jeroen disk 2
        hard_disk_list.Add("0025_388A_81BB_14B5.")   'Zweet kamer, Jeroen 

        user_list.Add("lennardh")
        hard_disk_list.Add("141190402709")       'VTK PC, Lennard Hubert

        user_list.Add("Peterdw")
        hard_disk_list.Add("134309552747")       'VTK PC, Peter de Wild

        user_list.Add("Jeffreyvdz")
        hard_disk_list.Add("ACE4_2E81_7006_2BD9.")     'VTK Laptop, Jeffrey van der Zwart

        user_list.Add("Twana")
        hard_disk_list.Add("ACE4_2E81_7006_2BD7.")     'VTK Laptop, Twan Akbheis

        user_list.Add("robru")
        hard_disk_list.Add("174741803447")      'VTK Laptop, Rob Ruiter

        nu = Now()
        nu2 = CDate("2019-12-01 00:00:00")
        life_time = CInt((nu2 - nu).TotalDays)
        Label101.Text = "Expire " & life_time.ToString

        TextBox28.Text = "Q" & Now.ToString("yy") & ".10"

        Pro_user = Environment.UserName     'User name on the screen
        HD_number = HardDisc_Id()           'Harddisk identification
        Me.Text &= "  (" & Pro_user & ")"

        'Check user name 
        For i = 0 To user_list.Count - 1
            If StrComp(LCase(Pro_user), LCase(user_list.Item(i))) = 0 Then pass_name = True
        Next

        'Check disc_id
        For i = 0 To hard_disk_list.Count - 1
            If CBool(HD_number = Trim(hard_disk_list(i))) Then pass_disc = True
        Next

        If pass_name = False Or pass_disc = False Then
            MessageBox.Show("VTK Cyclone selection program" & vbCrLf & "Access denied, contact GPa" & vbCrLf)
            MessageBox.Show("User_name= " & Pro_user & ", Pass name= " & pass_name.ToString)
            MessageBox.Show("HD_id= *" & HD_number & "*" & ", Pass disc= " & pass_disc.ToString)
            Environment.Exit(0)
        End If

        If life_time < 0 Then
            MessageBox.Show("Program lease Is Expired, contact GPa")
            Environment.Exit(0)
        End If

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        'De weerstandscoefficient inlet-air volgt uit het cycloon type
        weerstand_coef_air = {7, 7, 7, 7, 7.5, 9.5, 14.5}
        weerstand_coef_dust = {0, 7.927, 8.26, 7.615, 6.606, 6.175, 0}

        For hh = 0 To (cyl_dimensions.Length - 1)  'Fill combobox1 cyclone types
            words = cyl_dimensions(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 5                 'Select Cyclone type

        TextBox20.Text = "AA cyclone is a AC850 with diameter of 300mm" & vbCrLf
        TextBox20.Text &= "Load above 5 gr/m3 is considerde a high load" & vbCrLf
        TextBox20.Text &= "Cyclones can not choke" & vbCrLf

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles button1.Click, TabPage1.Enter, numericUpDown3.ValueChanged, numericUpDown2.ValueChanged, numericUpDown14.ValueChanged, NumericUpDown1.ValueChanged, numericUpDown5.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown18.ValueChanged, ComboBox1.SelectedIndexChanged, numericUpDown9.ValueChanged, numericUpDown8.ValueChanged, numericUpDown7.ValueChanged, numericUpDown6.ValueChanged, numericUpDown12.ValueChanged, numericUpDown11.ValueChanged, numericUpDown10.ValueChanged, numericUpDown13.ValueChanged, CheckBox2.CheckedChanged, NumericUpDown22.ValueChanged, NumericUpDown4.ValueChanged
        Dust_load_correction()
        Get_input_and_calc()
        Calc_loss_gvg()             'Calc according Guus
    End Sub
    Private Sub Get_input_and_calc()
        Dim words() As String
        Dim cyl_dim(20) As Double
        Dim db As Double            'Body diameter
        Dim Flow As Double
        Dim dp_inlet_gas As Double
        Dim dp_inlet_dust As Double
        Dim ro_gas, ro_particle, visco As Double
        Dim wc_air, wc_dust As Double
        Dim no_cycl As Double   'Number cyclones
        Dim stofb As Double
        Dim tot_kgh As Double       'Dust inlet per hour totaal 
        Dim kgh As Double           'Dust inlet per hour/cycloon 
        Dim kgs As Double           'Dust inlet per second

        Dim in_hoog As Double       '[m]
        Dim in_breed As Double      '[m]
        Dim dia_outlet As Double    '[m] gas outlet

        Dim inlet_velos As Double   '[m/s]
        Dim outlet_velos As Double  '[m/s]
        Dim total_input_weight As Double
        Dim h18, h19 As Double
        Dim j18, i18 As Double
        Dim l18, k19, k41 As Double

        If (ComboBox1.SelectedIndex > -1) Then     'Prevent exceptions
            words = cyl_dimensions(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            For hh = 1 To 15
                cyl_dim(hh) = CDbl(words(hh))   'Cyclone dimensions
            Next
            no_cycl = NumericUpDown20.Value     'Paralelle cyclonen
            db = numericUpDown5.Value / 1000    '[m] Body diameter
            in_hoog = cyl_dim(1) * db           '[m]
            in_breed = cyl_dim(2) * db          '[m]
            dia_outlet = cyl_dim(6) * db        '[m] 
            Flow = NumericUpDown1.Value / 3600  '[m3/s]
            Flow /= no_cycl                     '[m3/s/cycloon]
            ro_gas = numericUpDown3.Value       '[kg/m3]
            ro_particle = numericUpDown2.Value  '[kg/m3]
            visco = numericUpDown14.Value       '[cPoise]
            stofb = NumericUpDown4.Value        '[g/Am3]

            '----------- inlaat snelheid ---------------------
            inlet_velos = Flow / (in_breed * in_hoog)

            '----------- uitlaat snelheid ---------------------
            outlet_velos = Flow / ((PI / 4) * dia_outlet ^ 2)   '[m/s]

            '----------- Pressure loss cyclone----------------------
            wc_air = weerstand_coef_air(ComboBox1.SelectedIndex)
            wc_dust = weerstand_coef_dust(ComboBox1.SelectedIndex)
            dp_inlet_gas = 0.5 * ro_gas * inlet_velos ^ 2 * wc_air
            dp_inlet_dust = 0.5 * ro_gas * inlet_velos ^ 2 * wc_dust

            '----------- stof belasting ------------
            kgs = Flow * stofb / 1000               '[kg/s/cycloon]
            kgh = kgs * 3600                        '[kg/h/cycloon]
            tot_kgh = kgh * no_cycl                 '[kg/h] total

            '----------- K_stokes-----------------------------------
            _K_stokes = db * 2000 * visco * 16 / (ro_particle * 0.0181 * inlet_velos)
            _K_stokes = Sqrt(_K_stokes)

            '----------- presenteren ----------------------------------
            TextBox36.Text = Flow.ToString("0.000")                 '[m3/s] flow

            '----------- presenteren afmetingen ------------------------------
            TextBox1.Text = (in_hoog).ToString("0.000")              'inlaat breedte
            TextBox2.Text = (in_breed).ToString("0.000")             'Inlaat hoogte
            TextBox3.Text = (cyl_dim(3) * db).ToString("0.000")      'Inlaat lengte
            TextBox4.Text = (cyl_dim(4) * db).ToString("0.000")      'Inlaat hartmaat
            TextBox5.Text = (cyl_dim(5) * db).ToString("0.000")      'Inlaat afschuining
            TextBox6.Text = (cyl_dim(6) * db).ToString("0.000")      'Uitlaat keeldia inw.
            TextBox7.Text = (cyl_dim(7) * db).ToString("0.000")      'Uitlaat flensdiameter inw.
            TextBox8.Text = (cyl_dim(8) * db).ToString("0.000")      'Lengte insteekpijp inw.
            TextBox9.Text = (cyl_dim(9) * db).ToString("0.000")      'Lengte romp + conus
            TextBox10.Text = (cyl_dim(10) * db).ToString("0.000")    'Lengte romp
            TextBox11.Text = (cyl_dim(11) * db).ToString("0.000")    'Lengte çonus
            TextBox12.Text = (cyl_dim(12) * db).ToString("0.000")    'Dia_conus / 3P-pijp
            TextBox13.Text = (cyl_dim(13) * db).ToString("0.000")    'Lengte 3P-pijp
            TextBox14.Text = (cyl_dim(14) * db).ToString("0.000")    'Dia_conus / 3P-pijp
            TextBox15.Text = (cyl_dim(15) * db).ToString("0.000")    'Lengte 3P-pijp
            TextBox16.Text = inlet_velos.ToString("0.0")             'inlaat snelheid
            TextBox17.Text = dp_inlet_gas.ToString("0")              '[Pa] Pressure loss inlet-gas
            TextBox19.Text = (dp_inlet_gas / 100).ToString("0.0")    '[mbar] Pressure loss inlet-gas
            TextBox48.Text = dp_inlet_dust.ToString("0")             '[Pa]Pressure loss inlet-dust

            TextBox22.Text = outlet_velos.ToString("0.0")            'uitlaat snelheid
            TextBox23.Text = _K_stokes.ToString("0.000")             'Stokes waarde tov Standaard cycloon
            TextBox37.Text = numericUpDown5.Value.ToString           'Cycloone diameter
            TextBox38.Text = CType(ComboBox1.SelectedItem, String)   'Cycloon type

            Draw_chart1()
            Draw_chart2()
            '---------- Check speed ---------------
            If inlet_velos < 10 Or inlet_velos > 30 Then
                TextBox16.BackColor = Color.Red
            Else
                TextBox16.BackColor = Color.LightGreen
            End If

            '---------- Check dp ---------------
            If dp_inlet_gas > 2000 Then
                TextBox17.BackColor = Color.Red
                TextBox19.BackColor = Color.Red
            Else
                TextBox17.BackColor = Color.LightGreen
                TextBox19.BackColor = Color.LightGreen
            End If

            '--------- Inlet korrel-groep data -----------
            Init_groups()

            '--------- overall resultaat --------------------
            total_input_weight = 0

            DataGridView1.Columns(0).HeaderText = "Dia class"
            DataGridView1.Columns(1).HeaderText = "Feed psd cum"
            DataGridView1.Columns(2).HeaderText = "Feed psd diff"
            DataGridView1.Columns(3).HeaderText = "Loss % of feed"
            DataGridView1.Columns(4).HeaderText = "Loss abs [%]"
            DataGridView1.Columns(5).HeaderText = "Loss psd cum"
            DataGridView1.Columns(6).HeaderText = "Catch abs"
            DataGridView1.Columns(7).HeaderText = "Catch psd cum"
            DataGridView1.Columns(8).HeaderText = "Grade eff."

            For h = 0 To 22
                '---- diameter
                DataGridView1.Rows.Item(h).Cells(0).Value = guus(h * 5).d_ave.ToString("0.000")

                '---- feed
                DataGridView1.Rows.Item(h).Cells(1).Value = guus(h * 5).psd_cump.ToString("0.0000")
                If h > 0 Then
                    h18 = CDbl(DataGridView1.Rows.Item(h - 1).Cells(1).Value)
                Else
                    h18 = 100
                End If
                h19 = CDbl(DataGridView1.Rows.Item(h).Cells(1).Value)
                DataGridView1.Rows.Item(h).Cells(2).Value = (h18 - h19).ToString("0.000")

                '---- loss
                If CheckBox1.Checked Then
                    DataGridView1.Rows.Item(h).Cells(3).Value = (guus(h * 5).loss_overall * 100).ToString("0.0000")
                Else
                    DataGridView1.Rows.Item(h).Cells(3).Value = (guus(h * 5).loss_overall_C * 100).ToString("0.0000")
                End If
                i18 = CDbl(DataGridView1.Rows.Item(h).Cells(2).Value)
                j18 = CDbl(DataGridView1.Rows.Item(h).Cells(3).Value)
                DataGridView1.Rows.Item(h).Cells(4).Value = (i18 * j18 / 100).ToString("0.0000")
                If h > 0 Then
                    l18 = CDbl(DataGridView1.Rows.Item(h - 1).Cells(5).Value)
                Else
                    l18 = 100
                End If
                k19 = CDbl(DataGridView1.Rows.Item(h).Cells(4).Value)
                Double.TryParse(TextBox58.Text, k41)
                DataGridView1.Rows.Item(h).Cells(5).Value = (l18 - 100 * k19 / k41).ToString("0.0000")

                '---- catch
                If h > 0 Then
                    l18 = CDbl(DataGridView1.Rows.Item(h - 1).Cells(5).Value)
                Else
                    l18 = 100
                End If

                DataGridView1.Rows.Item(h).Cells(6).Value = "mm"
                DataGridView1.Rows.Item(h).Cells(7).Value = "kk"
                '---- efficiency
                DataGridView1.Rows.Item(h).Cells(8).Value = "ll"
            Next h
            DataGridView1.AutoResizeColumns()

            '---------- Calc diameter with 50% separation ---
            '---------- present -------
            TextBox42.Text = Calc_dia_particle(1.0).ToString("0.000")     '[mu] @ 100% loss
            TextBox26.Text = Calc_dia_particle(0.95).ToString("0.000")    '[mu] @  95% lost
            TextBox31.Text = Calc_dia_particle(0.9).ToString("0.000")     '[mu] @  90% lost
            TextBox32.Text = Calc_dia_particle(0.5).ToString("0.000")     '[mu] @  50% lost
            TextBox33.Text = Calc_dia_particle(0.1).ToString("0.000")     '[mu] @  10% lost
            TextBox41.Text = Calc_dia_particle(0.05).ToString("0.000")    '[mu] @   5% lost

            TextBox39.Text = kgh.ToString("0")          'Stof inlet
            TextBox40.Text = tot_kgh.ToString("0")      'Stof inlet totaal
        End If


    End Sub
    Private Sub Init_groups()
        DataGridView1.ColumnCount = 10
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(24)
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        '[mu] Class lower particle diameter limit diameter
        korrel_grp(0).dia_small = 0
        korrel_grp(1).dia_small = 10
        korrel_grp(2).dia_small = 15
        korrel_grp(3).dia_small = 20
        korrel_grp(4).dia_small = 30
        korrel_grp(5).dia_small = 40
        korrel_grp(6).dia_small = 50
        korrel_grp(7).dia_small = 60

        '[mu] Class upper particle diameter limit diameter
        korrel_grp(0).dia_big = 10
        korrel_grp(1).dia_big = 15
        korrel_grp(2).dia_big = 20
        korrel_grp(3).dia_big = 30
        korrel_grp(4).dia_big = 40
        korrel_grp(5).dia_big = 50
        korrel_grp(6).dia_big = 60
        korrel_grp(7).dia_big = 80

        korrel_grp(0).class_wght_cum_pro = numericUpDown6.Value / 100  'Percentale van de inlaat stof belasting
        korrel_grp(1).class_wght_cum_pro = numericUpDown7.Value / 100
        korrel_grp(2).class_wght_cum_pro = numericUpDown8.Value / 100
        korrel_grp(3).class_wght_cum_pro = numericUpDown9.Value / 100
        korrel_grp(4).class_wght_cum_pro = numericUpDown10.Value / 100
        korrel_grp(5).class_wght_cum_pro = numericUpDown11.Value / 100
        korrel_grp(6).class_wght_cum_pro = numericUpDown12.Value / 100
        korrel_grp(7).class_wght_cum_pro = numericUpDown13.Value / 100

    End Sub
    '-------- Bereken het verlies getal NIET gecorrigeerd -----------
    '----- de input is de GEMIDDELDE korrel grootte-----------
    Private Function Calc_verlies(korrel_g As Double, present As Boolean) As Double
        Dim words() As String
        Dim dia_Kcrit, fac_m, fac_a, fac_k As Double
        Dim verlies As Double = 1

        If (ComboBox1.SelectedIndex > -1) Then
            '-------------- korrelgrootte factoren ------
            words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))

            dia_Kcrit = CDbl(words(1))

            '-------- de grafieken zijn in 2 delen gesplits voor hogere nauwkeurigheid----------
            If korrel_g < dia_Kcrit Then
                fac_m = CDbl(words(2))
                fac_k = CDbl(words(3))
                fac_a = CDbl(words(4))
            Else
                fac_m = CDbl(words(5))
                fac_k = CDbl(words(6))
                fac_a = CDbl(words(7))
            End If

            If ((korrel_g / _K_stokes) - fac_m) > 0 Then
                verlies = (((korrel_g / _K_stokes) - fac_m) / fac_k) ^ fac_a
                verlies = Math.E ^ -verlies
            Else
                verlies = 1.0        '100% loss (very small particle)
            End If

        End If
        Return (verlies)
    End Function

    '-------- Bereken het verlies getal GECORRIGEERD -----------
    '----- de input is de GEMIDDELDE korrel grootte-----------
    Private Sub Calc_verlies_corrected(ByRef grp As GvG_Calc_struct)
        Dim words() As String
        Dim dia_Kcrit, fac_m, fac_a, fac_k As Double
        Dim cor1, cor2 As Double

        If (ComboBox1.SelectedIndex > -1) Then

            '-------------- korrelgrootte factoren ------
            words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            dia_Kcrit = CDbl(words(1))

            '-------- de grafieken zijn in 2 delen gesplits voor hogere nauwkeurigheid----------
            If grp.dia < dia_Kcrit Then
                fac_m = CDbl(words(2))
                fac_k = CDbl(words(3))
                fac_a = CDbl(words(4))
            Else
                fac_m = CDbl(words(5))
                fac_k = CDbl(words(6))
                fac_a = CDbl(words(7))
            End If

            '----- Insteek pijp corectie correctie -------
            cor1 = NumericUpDown22.Value    'Correctie insteek pijp

            'Hoge stof belasting correctie acc VT-UK
            Double.TryParse(TextBox47.Text, cor2)

            grp.loss_overall_C = grp.loss_overall ^ (cor1 * cor2)

        End If
        If fac_a > 0 Then
            TextBox24.Text &= "---------------" & grp.ToString & vbCrLf
            TextBox24.Text &= "cor1 = " & cor1.ToString & ", cor2 = " & cor2.ToString & vbCrLf
            TextBox24.Text &= "fac_m = " & fac_m.ToString & ", fac_k = " & fac_k.ToString & ", fac_a = " & fac_a.ToString & vbCrLf
            TextBox24.Text &= "grp.d_ave_K = " & grp.d_ave_K.ToString & vbCrLf
            TextBox24.Text &= "grp.loss_overall_C = " & grp.loss_overall_C.ToString & vbCrLf
            TextBox24.Text &= "guus(1).loss_overall_C. = " & guus(1).loss_overall_C.ToString & vbCrLf
        End If

    End Sub

    'Note dp(95) meaning with this diameter 95% is lost
    'Calculate the diameter at which qq% is lost
    Private Function Calc_dia_particle(qq As Double) As Double
        Dim dia_result As Double = 0
        Dim words() As String
        Dim dia_Kcrit As Double
        Dim d1, d2 As Double
        Dim cor1, cor2 As Double 'Insteek pijp
        Dim fac_m, fac_k, fac_a As Double

        If qq > 1 Then MessageBox.Show("Loss > 100% is impossible, Line 486, qq= " & qq.ToString)

        If (ComboBox1.SelectedIndex > -1) Then 'Prevent exceptions

            '----- Insteek pijp corectie correctie -------
            cor1 = NumericUpDown22.Value    'Correctie insteek pijp

            'Hoge stof belasting correctie acc VT-UK
            Double.TryParse(TextBox47.Text, cor2)

            '-------------- korrelgrootte factoren ------
            words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            dia_Kcrit = CDbl(words(1))   'Is in fact d/K(crit)

            '=$K$38*$B$51*((-LN(H56^(1/($B$54*$B$64))))^(1/$K$40))+$K$36*$B$51

            '---- diameter particle kleiner dan de diameter kritisch
            fac_m = CDbl(words(2))
            fac_k = CDbl(words(3))
            fac_a = CDbl(words(4))
            d1 = fac_k * _K_stokes * ((-Math.Log(qq ^ (1 / (cor1 * cor2))))) ^ (1 / fac_a) + fac_m * _K_stokes

            '---- diameter particle groter dan de diameter kritisch
            fac_m = CDbl(words(5))
            fac_k = CDbl(words(6))
            fac_a = CDbl(words(7))
            d2 = fac_k * _K_stokes * ((-Math.Log(qq ^ (1 / (cor1 * cor2))))) ^ (1 / fac_a) + fac_m * _K_stokes


            '=ALS(I56/$B$51<K$41    ,I56,J56)
            'I56= 87.2
            'B51= Kstokes= 1.25
            'K41= 7.8


            If ((d1 / _K_stokes) < dia_Kcrit) Then
                dia_result = d1     'diameter kleiner kritisch
            Else
                dia_result = d2     'diameter groter kritisch
            End If

            'TextBox24.Text &= "_K_stokes" & _K_stokes.ToString & vbCrLf
        End If
        Return (dia_result)
    End Function
    '---- According to VT-UK -----
    Private Sub Dust_load_correction()
        Dim f1, f2, f3, f4, f, f_used As Double
        Dim dst As Double

        dst = NumericUpDown4.Value / 1000 'Dust load dimension is [kg/Am3}

        f1 = 0.97833 + 2.918055 * dst - 39.3739 * dst ^ 2 + 472.0149 * dst ^ 3 - 769.586 * dst ^ 4
        f2 = -0.30338 + 21.91961 * dst - 73.5039 * dst ^ 2 + 112.485 * dst ^ 3 - 63.4408 * dst ^ 4
        f3 = 2.043212 + 0.725352 * dst - 0.2663 * dst ^ 2 + 0.04299 * dst ^ 3 - 0.00233 * dst ^ 4
        f4 = 2.853325 + 0.019026 * dst - 0.00036 * dst ^ 2 + 0.000003 * dst ^ 3 - 0.0000000065 * dst ^ 4

        Select Case dst
            Case < 0.02
                f = 1
            Case < 0.14
                f = f1
            Case < 0.7
                f = f2
            Case < 9.1
                f = f3
            Case Else
                f = f4
        End Select

        If (CheckBox2.Checked) Then
            f_used = f
        Else
            f_used = 1
        End If

        TextBox47.Text = f.ToString("0.000")
        TextBox55.Text = f_used.ToString("0.000")
    End Sub

    Private Sub Draw_chart1()
        '-------
        Dim s_points(100, 2) As Double
        Dim h As Integer
        Dim sdia As Integer

        Chart1.Series.Clear()
        Chart1.ChartAreas.Clear()
        Chart1.Titles.Clear()
        Chart1.ChartAreas.Add("ChartArea0")

        Chart1.Series.Add("Series" & h.ToString)
        Chart1.Series(h).ChartArea = "ChartArea0"
        Chart1.Series(h).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart1.Series(h).BorderWidth = 2
        Chart1.Series(h).IsVisibleInLegend = False

        Chart1.Titles.Add("Loss Curve")
        Chart1.ChartAreas("ChartArea0").AxisX.Title = "particle dia [mu]"

        Chart1.ChartAreas("ChartArea0").AxisY.Title = "Loss [%] (niet gevangen)"
        Chart1.ChartAreas("ChartArea0").AxisY.Minimum = 0       'Loss
        Chart1.ChartAreas("ChartArea0").AxisY.Maximum = 100     'Loss
        Chart1.ChartAreas("ChartArea0").AxisY.Interval = 10     'Interval
        Chart1.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
        Chart1.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True
        Chart1.ChartAreas("ChartArea0").AxisX.MinorGrid.Enabled = True
        Chart1.ChartAreas("ChartArea0").AxisY.MinorGrid.Enabled = True

        If CheckBox1.Checked Then
            Chart1.ChartAreas("ChartArea0").AxisX.IsLogarithmic = True
            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 1     'Particle size
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = 100   'Particle size
        Else
            Chart1.ChartAreas("ChartArea0").AxisX.IsLogarithmic = False
            Chart1.ChartAreas("ChartArea0").AxisX.Minimum = 0     'Particle size
            Chart1.ChartAreas("ChartArea0").AxisX.Maximum = 20    'Particle size
        End If

        '----- now calc chart points --------------------------
        Integer.TryParse(TextBox42.Text, sdia)
        s_points(0, 0) = sdia   'Particle diameter [mu]
        s_points(0, 1) = 100    '100% loss
        For h = 1 To 40
            s_points(h, 0) = h                                   'Particle diameter [mu]
            s_points(h, 1) = Calc_verlies(s_points(h, 0), False) * 100  'Loss [%]
        Next

        '------ now present-------------
        For h = 0 To 40 - 1   'Fill line chart
            Chart1.Series(0).Points.AddXY(s_points(h, 0), s_points(h, 1))
        Next h
    End Sub
    Private Sub Draw_chart2()
        Dim s_points(100, 2) As Double
        Dim h As Integer
        Dim sdia As Integer

        Chart2.Series.Clear()
        Chart2.ChartAreas.Clear()
        Chart2.Titles.Clear()
        Chart2.ChartAreas.Add("ChartArea0")

        Chart2.Series.Add("Series" & h.ToString)
        Chart2.Series(h).ChartArea = "ChartArea0"
        Chart2.Series(h).ChartType = DataVisualization.Charting.SeriesChartType.Line
        Chart2.Series(h).BorderWidth = 2
        Chart2.Series(h).IsVisibleInLegend = False

        Chart2.Titles.Add("Loss Curve")
        Chart2.ChartAreas("ChartArea0").AxisX.Title = "particle dia [mu]"
        Chart2.ChartAreas("ChartArea0").AxisY.Minimum = 0     'Loss
        Chart2.ChartAreas("ChartArea0").AxisY.Maximum = 100   'Loss
        Chart2.ChartAreas("ChartArea0").AxisX.Minimum = 0     'Particle size
        Chart2.ChartAreas("ChartArea0").AxisX.Maximum = 20    'Particle size

        '----- now calc chart poins --------------------------
        Integer.TryParse(TextBox42.Text, sdia)
        s_points(0, 0) = sdia   'Particle diameter [mu]
        s_points(0, 1) = 100    '100% loss
        For h = 1 To 40
            s_points(h, 0) = h                                   'Particle diameter [mu]
            s_points(h, 1) = Calc_verlies(s_points(h, 0), False) * 100  'Loss [%]
        Next

        '------ now present-------------
        For h = 0 To 40 - 1   'Fill line chart
            Chart2.Series(0).Points.AddXY(s_points(h, 0), s_points(h, 1))
        Next h
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage9.Enter, CheckBox1.CheckedChanged

        Calc_loss_gvg()             'Calc according Guus
        Present_loss_grid()         'Present the results
        Draw_chart1()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox28.Text.Trim.Length > 0 And TextBox29.Text.Trim.Length > 0 Then
            Save_tofile()
        Else
            MessageBox.Show("Complete Quote and Tag number")
        End If
    End Sub
    Private Sub Save_tofile()

        Dim temp_string, user As String

        user = Trim(Environment.UserName)         'User name on the screen
        Dim filename As String = "Cyclone_select_" & TextBox28.Text & "_" & TextBox29.Text & DateTime.Now.ToString("_yyyy_MM_dd_") & user & ".vtk2"
        filename = Replace(filename, Chr(32), Chr(95)) 'Replace the space's

        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox29.Text) Then
            TextBox29.Text = "-"
        End If

        '-------- Project information -----------------
        temp_string = TextBox28.Text & ";" & TextBox29.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all numeric controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        all_num = all_num.OrderBy(Function(x) x.Name).ToList()      'Alphabetical order
        For i = 0 To all_num.Count - 1
            Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
            temp_string &= grbx.Value.ToString & ";"
            TextBox24.Text &= grbx.Name.ToString & "value= " & grbx.Value.ToString & vbTab & " is Saved to file" & vbCrLf
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save ----------------
        FindControlRecursive(all_combo, Me, GetType(ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
            temp_string &= grbx.SelectedItem.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all checkbox controls and save --------------------
        FindControlRecursive(all_check, Me, GetType(CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As CheckBox = CType(all_check(i), CheckBox)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all radio controls and save ---------------
        FindControlRecursive(all_radio, Me, GetType(RadioButton))   'Find the control
        all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_radio.Count - 1
            Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_tmp)) Then System.IO.Directory.CreateDirectory(dirpath_tmp)
        Catch ex As Exception
            MessageBox.Show("Create directory without VTK intranet (L578)" & vbCrLf & ex.Message)
        End Try

        Try
            If (Not System.IO.Directory.Exists(dirpath_Temp)) Then System.IO.Directory.CreateDirectory(dirpath_Temp)
            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
        Catch ex As Exception
            MessageBox.Show("Can not create directory on the VTK intranet (L6286) " & vbCrLf & vbCrLf & ex.Message)
        End Try

        Try
            If CInt(temp_string.Length.ToString) > 100 Then      'String may be empty
                If Directory.Exists(dirpath_Eng) Then
                    File.WriteAllText(dirpath_Eng & filename, temp_string, Encoding.ASCII)     'used at VTK with intranet
                Else
                    File.WriteAllText(dirpath_tmp & filename, temp_string, Encoding.ASCII)     'used at home
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Line 6298, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    'Retrieve control settings from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Private Sub Read_file()
        Dim control_words(), words() As String
        Dim i As Integer
        Dim ttt As Double
        Dim k As Integer = 0
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        OpenFileDialog1.FileName = "Cyclone_select_*"

        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_tmp  'used at home
        End If

        OpenFileDialog1.Title = "Open a Text File"
        OpenFileDialog1.Filter = "VTK2 Files|*.vtk2|VTK1 file|*.vtk"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim readText As String = File.ReadAllText(OpenFileDialog1.FileName, Encoding.ASCII)
            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content
            '----- retrieve Project information ----------------------
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split the read file content
            TextBox28.Text = words(0)                  'Project number
            TextBox29.Text = words(1)                  'Tag no

            '---------- Retrieve Numeric controls from disk-----------------
            FindControlRecursive(all_num, Me, GetType(NumericUpDown))               'Find the numericupdowns
            all_num = all_num.OrderBy(Function(x) x.Name).ToList()                  'Sort in Alphabetical order
            words = control_words(1).Split(separators, StringSplitOptions.None)     'Split the read file content
            For i = 0 To all_num.Count - 1
                Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal --
                If (i < words.Length - 1) Then
                    If Not (Double.TryParse(words(i + 1), ttt)) Then
                        MessageBox.Show("Numeric controls conversion problem occured")
                        TextBox24.Text &= grbx.Name.ToString & ", ttt= |" & ttt.ToString & "| Cannot be converted to number" & vbCrLf
                    End If

                    If ttt <= grbx.Maximum And ttt >= grbx.Minimum Then
                        grbx.Value = CDec(ttt)          'OK
                    Else
                        TextBox24.Text &= grbx.Name.ToString & " ttt= " & ttt.ToString & " range=, " & grbx.Minimum & "-" & grbx.Maximum & " Minimum value Is used " & vbCrLf
                        grbx.Value = grbx.Minimum       'NOK
                    End If
                Else
                    TextBox24.Text &= "Warning last Numeric-Updown-controls Not found In file" & vbCrLf
                End If
            Next

            '---------- Retrieve  combobox controls -----------------
            FindControlRecursive(all_combo, Me, GetType(ComboBox))
            all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()          'Alphabetical order
            words = control_words(2).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_combo.Count - 1
                Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    grbx.SelectedItem = words(i + 1)
                Else
                    TextBox24.Text &= "Warning last combobox Not found In file" & vbCrLf
                End If
            Next

            '---------- Retrieve  checkbox controls -----------------
            FindControlRecursive(all_check, Me, GetType(CheckBox))
            all_check = all_check.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(3).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_check.Count - 1
                Dim grbx As CheckBox = CType(all_check(i), CheckBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    TextBox24.Text &= "Warning last checkbox Not found In file" & vbCrLf
                End If
            Next

            '---------- Retrieve  radiobuttons controls -----------------
            FindControlRecursive(all_radio, Me, GetType(RadioButton))
            all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(4).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_radio.Count - 1
                Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal radiobuttons--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    TextBox24.Text &= "Warning last radiobutton Not found In file" & vbCrLf
                End If
            Next
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Read_file()
    End Sub
    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

    Public Function HardDisc_Id() As String
        'Add system.management as reference !!
        'imports system.management
        Dim tmpStr2 As String = ""
        Dim myScop As New ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
        Dim oQuer As New SelectQuery("SELECT * FROM WIN32_DiskDrive")

        Dim oResult As New ManagementObjectSearcher(myScop, oQuer)
        Dim oIte As ManagementObject
        Dim oPropert As PropertyData
        For Each oIte In oResult.Get()
            For Each oPropert In oIte.Properties
                If Not oPropert.Value Is Nothing AndAlso oPropert.Name = "SerialNumber" Then
                    tmpStr2 = oPropert.Value.ToString
                    Exit For
                End If
            Next
            Exit For
        Next
        Return (Trim(tmpStr2))         'Harddisk identification
    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox28.Text.Trim.Length > 0 Then
            Get_input_and_calc()
            Write_to_word_com() 'Commercial data to Word
        Else
            MessageBox.Show("Enter Quote nummer and Tag, then Export sizing data to Word")
        End If
    End Sub

    'Write COMMERCIAL data to Word 
    'see https://msdn.microsoft.com/en-us/library/office/aa192495(v=office.11).aspx
    Private Sub Write_to_word_com()
        Dim bmp_tab_page1 As New Bitmap(TabPage1.Width, TabPage1.Height)
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara4 As Word.Paragraph

        Dim chart_size As Integer = 140  '% of original picture size
        Dim file_name As String
        Dim row As Integer = 0
        Try
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            oDoc.PageSetup.LeftMargin = 60
            oDoc.PageSetup.TopMargin = 35
            oDoc.PageSetup.BottomMargin = 10
            oDoc.PageSetup.RightMargin = 20
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait
            oDoc.PageSetup.PaperSize = Word.WdPaperSize.wdPaperA4
            'oDoc.PageSetup.VerticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalCenter

            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "VTK Sales"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = 14
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 0.5                '24 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            '---------------Inlet data-------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 15, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 11
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 2
            oTable.Cell(row, 1).Range.Text = "Project number"
            oTable.Cell(row, 2).Range.Text = TextBox28.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Tag nummer "
            oTable.Cell(row, 2).Range.Text = TextBox29.Text

            row += 1
            oTable.Cell(row, 1).Range.Text = "Print date"
            oTable.Cell(row, 2).Range.Text = Now().ToString("MM-dd-yyyy")
            row += 2
            oTable.Cell(row, 1).Range.Text = "Flow"
            oTable.Cell(row, 2).Range.Text = NumericUpDown1.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[Am3/hr]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Temperature"
            oTable.Cell(row, 2).Range.Text = NumericUpDown18.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[c]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Druk"
            oTable.Cell(row, 2).Range.Text = (NumericUpDown19.Value / 100).ToString("0.0")
            oTable.Cell(row, 3).Range.Text = "[mbar]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Particle density "
            oTable.Cell(row, 2).Range.Text = numericUpDown2.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Gas density "
            oTable.Cell(row, 2).Range.Text = numericUpDown3.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Air viscosity"
            oTable.Cell(row, 2).Range.Text = numericUpDown14.Value.ToString("0.0000")
            oTable.Cell(row, 3).Range.Text = "[centi Poise]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Dust load"
            oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[gr/Am3]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Dust load (1 cyclone)"
            oTable.Cell(row, 2).Range.Text = TextBox39.Text
            oTable.Cell(row, 3).Range.Text = "[kg/hr]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "dp(50) "
            oTable.Cell(row, 2).Range.Text = TextBox32.Text
            oTable.Cell(row, 3).Range.Text = "[mu]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Emssion"
            oTable.Cell(row, 2).Range.Text = TextBox18.Text
            oTable.Cell(row, 3).Range.Text = "[g/Am3]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(1)
            oTable.Columns(3).Width = oWord.InchesToPoints(2)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------cyclone data-------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Cyclone data"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Cyclone type "
            oTable.Cell(row, 2).Range.Text = ComboBox1.SelectedItem.ToString
            row += 1
            oTable.Cell(row, 1).Range.Text = "Body diameter"
            oTable.Cell(row, 2).Range.Text = numericUpDown5.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "No parallel"
            oTable.Cell(row, 2).Range.Text = NumericUpDown20.Value.ToString

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(1)
            oTable.Columns(3).Width = oWord.InchesToPoints(2)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------Process data-------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 4, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Process data"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inlet speed "
            oTable.Cell(row, 2).Range.Text = TextBox16.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Outlet data"
            oTable.Cell(row, 2).Range.Text = TextBox22.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Pressure loss"
            oTable.Cell(row, 2).Range.Text = TextBox17.Text
            oTable.Cell(row, 3).Range.Text = "[Pa]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(1)
            oTable.Columns(3).Width = oWord.InchesToPoints(2)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------Calculation date-------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 11, 8)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Calc.data"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Lower dia [mu]"
            oTable.Cell(row, 2).Range.Text = "Upper dia [mu]"
            oTable.Cell(row, 3).Range.Text = "Average dia [mu]"
            oTable.Cell(row, 4).Range.Text = "Wght [kg/hr]"
            oTable.Cell(row, 5).Range.Text = "Wght [%]"
            oTable.Cell(row, 6).Range.Text = "Loss [%]"
            oTable.Cell(row, 7).Range.Text = "Loss [kg/hr]"

            For j = 0 To 8
                row += 1
                oTable.Cell(row, 1).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(0).Value, String)
                oTable.Cell(row, 2).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(1).Value, String)
                oTable.Cell(row, 3).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(2).Value, String)
                oTable.Cell(row, 4).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(3).Value, String)
                oTable.Cell(row, 5).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(4).Value, String)
                oTable.Cell(row, 6).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(5).Value, String)
                oTable.Cell(row, 7).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(6).Value, String)
            Next

            For j = 1 To 8
                oTable.Columns(j).Width = oWord.InchesToPoints(0.8)   'Change width of columns 
            Next
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------save Chart1 (Loss curve)---------------- 
            Draw_chart2()
            file_name = dirpath_Temp & "Chart_loss.Jpeg"
            Chart2.SaveImage(file_name, System.Drawing.Imaging.ImageFormat.Jpeg)
            oPara4 = oDoc.Content.Paragraphs.Add
            oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
            oPara4.Range.InlineShapes.AddPicture(file_name)
            oPara4.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
            oPara4.Range.InlineShapes.Item(1).ScaleWidth = chart_size       'Size
            oPara4.Range.InsertParagraphAfter()

        Catch ex As Exception
            MessageBox.Show(ex.Message & " Problem writing to Commercial data to Word ")  ' Show the exception's message.
        End Try
    End Sub
    'Air viscosity

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown21.ValueChanged
        Dim Visco As Double
        Visco = Air_visco(CDbl(NumericUpDown21.Value))
        TextBox30.Text = Visco.ToString("0.00000")
    End Sub

    'http://www-mdp.eng.cam.ac.uk/web/library/enginfo/aerothermal_dvd_only/aero/fprops/propsoffluids/node5.html
    'Sutherland Equation (Range -70 to 1600 celsius)
    Private Function Air_visco(temp As Double) As Double
        Dim C1, C2 As Double
        Dim vis As Double

        temp += 273.15  '[Celsius]-->[K]
        C1 = 1.458 * 10 ^ -5
        C2 = 110.4
        vis = C1 * temp ^ 1.5 / (temp + C2)
        Return (vis * 100)    '[kg/m-s]-->[centi Poise]
    End Function

    Private Sub Present_loss_grid()
        Dim j As Integer
        Dim total_abs_loss_C As Double = 0
        Dim total_abs_loss As Double = 0
        Dim total_psd_diff As Double = 0

        DataGridView2.ColumnCount = 18
        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add(111)
        'DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        DataGridView2.Columns(0).HeaderText = "Dia class [mu]"
        DataGridView2.Columns(1).HeaderText = "Dia average"
        DataGridView2.Columns(2).HeaderText = "Dia/k"
        DataGridView2.Columns(3).HeaderText = "Loss overall"
        DataGridView2.Columns(4).HeaderText = "Loss overall Corrected"
        DataGridView2.Columns(5).HeaderText = "Catch chart"     '
        DataGridView2.Columns(6).HeaderText = "Groep nummer"    '
        DataGridView2.Columns(7).HeaderText = "d1 lower dia"    '
        DataGridView2.Columns(8).HeaderText = "d2 upper dia"    '
        DataGridView2.Columns(9).HeaderText = "p1 input"        '
        DataGridView2.Columns(10).HeaderText = "p2 input"       '
        DataGridView2.Columns(11).HeaderText = "k"              '    
        DataGridView2.Columns(12).HeaderText = "m"              '

        DataGridView2.Columns(13).HeaderText = "i_psd cum"      '
        DataGridView2.Columns(14).HeaderText = "psd cum [%]"    '
        DataGridView2.Columns(15).HeaderText = "psd diff"       '
        DataGridView2.Columns(16).HeaderText = "loss abs [%]"   '
        DataGridView2.Columns(17).HeaderText = "loss corr abs [%]" '

        For row = 1 To 110  'Fill the DataGrid
            j = row - 1
            DataGridView2.Rows.Item(j).Cells(0).Value = guus(j).dia.ToString
            DataGridView2.Rows.Item(j).Cells(1).Value = guus(j).d_ave.ToString      'Average diameter
            DataGridView2.Rows.Item(j).Cells(2).Value = guus(j).d_ave_K.ToString    'Average dia/K stokes
            DataGridView2.Rows.Item(j).Cells(3).Value = guus(j).loss_overall.ToString   'Loss 
            DataGridView2.Rows.Item(j).Cells(4).Value = guus(j).loss_overall_C.ToString 'Loss 
            DataGridView2.Rows.Item(j).Cells(5).Value = guus(j).catch_chart.ToString    'Catch
            DataGridView2.Rows.Item(j).Cells(6).Value = guus(j).i_grp.ToString      'Groep nummer
            DataGridView2.Rows.Item(j).Cells(7).Value = guus(j).i_d1.ToString       'class lower dia limit
            DataGridView2.Rows.Item(j).Cells(8).Value = guus(j).i_d2.ToString       'class upper dia limit
            DataGridView2.Rows.Item(j).Cells(9).Value = guus(j).i_p1.ToString       'User input percentage
            DataGridView2.Rows.Item(j).Cells(10).Value = guus(j).i_p2.ToString      '
            DataGridView2.Rows.Item(j).Cells(11).Value = guus(j).i_k.ToString       'User input percentage
            DataGridView2.Rows.Item(j).Cells(12).Value = guus(j).i_m.ToString       '
            DataGridView2.Rows.Item(j).Cells(13).Value = guus(j).psd_cum.ToString   '
            DataGridView2.Rows.Item(j).Cells(14).Value = guus(j).psd_cump.ToString("0.0")   '[%]
            DataGridView2.Rows.Item(j).Cells(15).Value = guus(j).psd_dif.ToString("E3")     '[%]
            DataGridView2.Rows.Item(j).Cells(16).Value = guus(j).loss_abs.ToString("E3")    '[%]
            DataGridView2.Rows.Item(j).Cells(17).Value = guus(j).loss_abs_C.ToString("E3")  '[%]
            total_psd_diff += guus(j).psd_dif
            total_abs_loss += guus(j).loss_abs
            total_abs_loss_C += guus(j).loss_abs_C
        Next
        DataGridView2.Rows.Item(111).Cells(15).Value = total_psd_diff.ToString
        DataGridView2.Rows.Item(111).Cells(16).Value = total_abs_loss.ToString
        DataGridView2.Rows.Item(111).Cells(17).Value = total_abs_loss_C.ToString
    End Sub
    Private Sub Calc_loss_gvg()
        'This is the standard VTK cyclone calculation 
        Dim i As Integer = 0
        Dim dia_max As Double         'Above this diameter everything is caught
        Dim dia_min As Double       'Below this diameter nothing is caught
        Dim istep As Double         'Particle diameter step
        Dim sum_loss As Double
        Dim sum_loss_C As Double
        Dim sum_psd_diff As Double
        Dim loss_total As Double
        Dim perc_smallest_part As Double
        Dim fac_m As Double
        Dim words() As String

        guus(i).dia = Calc_dia_particle(1.0)
        guus(i).d_ave = guus(0).dia / 2                                 'Average diameter
        guus(i).d_ave_K = guus(0).d_ave / _K_stokes                     'dia/k_stokes
        guus(i).loss_overall = Calc_verlies(guus(0).d_ave_K, False)     '[-] loss overall
        Calc_verlies_corrected(guus(0))                                 '[-] loss overall corrected
        guus(i).catch_chart = (1 - guus(i).loss_overall_C) * 100        '[%]
        guus(i).i_grp = Find_class_limits(guus(i).dia, 5)               'groepnummer
        guus(i).i_d1 = Find_class_limits(guus(i).dia, 1)                'Lower limit diameter
        guus(i).i_d2 = Find_class_limits(guus(i).dia, 2)                'Upper limit diameter
        guus(i).i_p1 = Find_class_limits(guus(i).dia, 3)                'Percentage
        guus(i).i_p2 = Find_class_limits(guus(i).dia, 4)                'Percentage

        guus(i).i_k = Log(Math.Log(guus(i).i_p1) / Math.Log(guus(i).i_p2))
        guus(i).i_k /= Log(guus(i).i_d1 / guus(i).i_d2)
        guus(i).i_m = guus(i).i_d1 / ((-Log(guus(i).i_p1)) ^ (1 / guus(i).i_k))
        guus(i).psd_cum = Math.E ^ (-((guus(i).dia / guus(i).i_m) ^ guus(i).i_k))
        guus(i).psd_cump = guus(i).psd_cum * 100
        guus(i).psd_dif = 100 * (1 - guus(i).psd_cum)
        guus(i).loss_abs = guus(i).loss_overall * guus(i).psd_dif
        guus(i).loss_abs_C = guus(i).loss_overall_C * guus(i).psd_dif

        sum_psd_diff = guus(i).psd_dif
        sum_loss = guus(i).loss_abs
        sum_loss_C = guus(i).loss_abs_C

        '------ increment step --------
        'stapgrootte bij 110-staps logaritmische verdeling van het
        'deeltjesdiameter-bereik van loss=100% tot 0,00000001%
        'Deze wordt gebruikt voor het opstellen van de gefractioneerde
        'verliescurve.

        '-------------- korrelgrootte factoren ------
        If ComboBox1.SelectedIndex > -1 Then
            words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            '---- diameter kleiner dan dia kritisch
            fac_m = CDbl(words(2))
        End If

        perc_smallest_part = 0.0000001                      'smallest particle [%]
        dia_max = Calc_dia_particle(perc_smallest_part)     '=100% loss (biggest particle)
        dia_min = _K_stokes * fac_m                 'diameter smallest particle caught
        istep = (dia_max / dia_min) ^ (1 / 110)     'Calculation step

        'TextBox24.Text &= "dia_min = " & dia_min.ToString & vbCrLf
        'TextBox24.Text &= "dia_max = " & dia_max.ToString & vbCrLf
        'TextBox24.Text &= "istep = " & istep.ToString & vbCrLf

        For i = 1 To 110
            guus(i).dia = guus(i - 1).dia * istep
            guus(i).d_ave = (guus(i - 1).dia + guus(i).dia) / 2         'Average diameter
            guus(i).d_ave_K = guus(i).d_ave / _K_stokes                 'dia/k_stokes
            guus(i).loss_overall = Calc_verlies(guus(i).d_ave, False)   '[-] loss overall
            Calc_verlies_corrected(guus(i))                             '[-] loss overall corrected
            If CheckBox2.Checked Then
                guus(i).catch_chart = (1 - guus(i).loss_overall_C) * 100  '[%] Corrected
            Else
                guus(i).catch_chart = (1 - guus(i).loss_overall) * 100    '[%] NOT corrected
            End If
            guus(i).i_grp = Find_class_limits(guus(i).dia, 5)         'groepnummer
            guus(i).i_d1 = Find_class_limits(guus(i).dia, 1)          'Lower diameter limit
            guus(i).i_d2 = Find_class_limits(guus(i).dia, 2)          'Upper diameter limit
            guus(i).i_p1 = Find_class_limits(guus(i).dia, 3)          'Percentage
            guus(i).i_p2 = Find_class_limits(guus(i).dia, 4)          'Percentage
            If guus(i).i_p2 > 0.001 Then 'to prevent silly results
                guus(i).i_k = Log(Log(guus(i).i_p1) / Log(guus(i).i_p2)) / Log(guus(i).i_d1 / guus(i).i_d2)
                guus(i).i_m = guus(i).i_d1 / ((-Log(guus(i).i_p1)) ^ (1 / guus(i).i_k))
                guus(i).psd_cum = Math.E ^ (-((guus(i).dia / guus(i).i_m) ^ guus(i).i_k))
                guus(i).psd_cump = guus(i).psd_cum * 100
                guus(i).psd_dif = 100 * (guus(i - 1).psd_cum - guus(i).psd_cum)
            Else
                guus(i).i_k = 0
                guus(i).i_m = 0
                guus(i).psd_cum = 0
                guus(i).psd_cump = 0
                guus(i).psd_dif = 0
            End If
            guus(i).loss_abs = guus(i).loss_overall * guus(i).psd_dif
            guus(i).loss_abs_C = guus(i).loss_overall_C * guus(i).psd_dif
            '----- sum value -----
            sum_psd_diff += guus(i).psd_dif
            sum_loss += guus(i).loss_abs
            sum_loss_C += guus(i).loss_abs_C
        Next
        loss_total = sum_loss_C + ((100 - sum_psd_diff) * perc_smallest_part)

        '----------- present -----------
        TextBox56.Text = ComboBox1.Text
        TextBox57.Text = CheckBox2.Checked.ToString

        If CheckBox2.Checked Then
            TextBox58.Text = loss_total.ToString("0.00000")    'Corrected
            TextBox59.Text = (100 - loss_total).ToString("0.000")
            TextBox21.Text = TextBox59.Text
            TextBox60.Text = (NumericUpDown4.Value * loss_total / 100).ToString("0.000")
            TextBox18.Text = TextBox60.Text
        Else
            TextBox58.Text = sum_loss.ToString("0.00000")      'NOT Corrected
            TextBox59.Text = (100 - sum_loss).ToString("0.000")
            TextBox21.Text = TextBox59.Text
            TextBox60.Text = (NumericUpDown4.Value * sum_loss / 100).ToString("0.000")
            TextBox18.Text = TextBox60.Text
        End If
    End Sub
    'Determine the particle diameter class upper and lower limits
    Private Function Find_class_limits(dia As Double, noi As Integer) As Double
        Dim d1 As Double 'particle diameter [mu]
        Dim d2 As Double 'particle diameter [mu]
        Dim input_p1, input_p2 As Double
        Dim grp As Double 'groepnummer

        Dim ret As Double
        Select Case True
            Case dia < 10
                d1 = 10
                d2 = 15
                input_p1 = numericUpDown6.Value / 100
                input_p2 = numericUpDown7.Value / 100
                grp = 0
            Case dia >= 10 And dia < 15
                d1 = 10
                d2 = 15
                input_p1 = numericUpDown6.Value / 100
                input_p2 = numericUpDown7.Value / 100
                grp = 1
            Case dia >= 15 And dia < 20
                d1 = 15
                d2 = 20
                input_p1 = numericUpDown7.Value / 100
                input_p2 = numericUpDown8.Value / 100
                grp = 2
            Case dia >= 20 And dia < 30
                d1 = 20
                d2 = 30
                input_p1 = numericUpDown8.Value / 100
                input_p2 = numericUpDown9.Value / 100
                grp = 3
            Case dia >= 30 And dia < 40
                d1 = 30
                d2 = 40
                input_p1 = numericUpDown9.Value / 100
                input_p2 = numericUpDown10.Value / 100
                grp = 4
            Case dia >= 40 And dia < 50
                d1 = 40
                d2 = 50
                input_p1 = numericUpDown10.Value / 100
                input_p2 = numericUpDown11.Value / 100
                grp = 5
            Case dia >= 50 And dia < 60
                d1 = 50
                d2 = 60
                input_p1 = numericUpDown11.Value / 100
                input_p2 = numericUpDown12.Value / 100
                grp = 6
            Case dia >= 60 And dia < 80
                d1 = 60
                d2 = 80
                input_p1 = numericUpDown12.Value / 100
                input_p2 = numericUpDown13.Value / 100
                grp = 7
            Case Else
                d1 = 80
                d2 = 120
                input_p1 = 0.0001
                input_p2 = 0.00001
                grp = 8
        End Select

        '------ select the return variable -------------- 
        Select Case noi
            Case 1
                ret = d1
            Case 2
                ret = d2
            Case 3
                ret = input_p1
            Case 4
                ret = input_p2
            Case 5
                ret = grp
        End Select

        Return (ret)
    End Function

End Class
