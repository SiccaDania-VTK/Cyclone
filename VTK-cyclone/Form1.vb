Imports System.Globalization
Imports System.IO
Imports System.Management
Imports System.Math
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports VTK_cyclone
Imports Word = Microsoft.Office.Interop.Word
'------- Input data------
'This structure is required for the different operating cases of a cyclone
'Therefore the struct does only contain  the input information
'If the calculation is modified the new result will be found 
<Serializable()> Public Structure Input_struct
    Public case_name As String      'The case name
    Public FlowT As Double          'Air flow [Am3/h]
    Public dia_big() As Double      'Particle diameter inlet [mu]
    Public class_load() As Double   'group_weight_cum in de inlaat stroom [% weight]
    Public ro_gas As Double         '[kg/hr] Density 
    Public ro_solid As Double       '[kg/hr] Density 
    Public visco As Double          '[Centi Poise] Visco in 
    Public Temp As Double           '[c] Temperature 

    '===== stage #1 parameter ======
    Public Flow1 As Double          '[Am3/s] Air flow per cyclone 
    Public stofb1 As Double         '[g/Am3] Dust load inlet 
    Public emmis1 As Double         '[g/Am3] Dust emission 
    Public Efficiency1 As Double    '[%] Efficiency Stage #1 
    Public sum_loss1 As Double      '[-]Passed trough cyclone 
    Public sum_loss_C1 As Double    '[-] Passed trough cyclone Corrected
    Public loss_total1 As Double
    Public sum_psd_diff1 As Double
    Public Druk1 As Double          '[mbar] druk
    Public Ct1 As Integer           '[-] Cyclone type (eg AC435)
    Public Noc1 As Integer          '[-] Number paralle Cyclones
    Public db1 As Double            '[m] Diameter cyclone body
    Public inh1 As Double           '[m] inlet hoogte
    Public inb1 As Double           '[m] inlet breedte
    Public dout1 As Double          '[m] diameter zakbuis
    Public inv1 As Double           '[m/s] Inlet velocity cyclone
    Public outv1 As Double          '[m/s] Outlet velocity cyclone
    Public Kstokes1 As Double       'Stokes getal
    Public dpgas1 As Double         '[Pa] pressure loss gas
    Public dpdust1 As Double        '[Pa] pressure loss dust
    Public m1 As Double             'm factor loss curve d< dia critical
    Public stage1() As GvG_Calc_struct   'tbv calculatie stage #1
    Public Dmin1 As Double          'Smallest particle in calculation
    Public Dmax1 As Double          'Biggest particle in calculation


    '===== stage #2 parameters ======
    Public Flow2 As Double          '[Am3/s] Air flow per cyclone 
    Public stofb2 As Double         '[g/Am3] Dust load inlet 
    Public emmis2 As Double         '[g/Am3] Dust emission 
    Public sum_loss2 As Double      'Passed trough cyclone
    Public sum_loss_C2 As Double    'Passed trough cyclone Corrected
    Public loss_total2 As Double
    Public sum_psd_diff2 As Double
    Public Efficiency2 As Double    'Efficiency Stage #1 [%}
    Public Druk2 As Double          '[mbar] druk
    Public Ct2 As Integer           '[-] Cyclone type (eg AC435)
    Public Noc2 As Integer          '[-] Number paralle Cyclones
    Public db2 As Double            '[m] Diameter cyclone body
    Public inh2 As Double           '[m] inlet hoogte
    Public inb2 As Double           '[m] inlet breedte
    Public dout2 As Double          '[m] diameter zakbuis
    Public inv2 As Double           '[m/s] Inlet velocity cyclone
    Public outv2 As Double          '[m/s] Outlet velocity cyclone
    Public Kstokes2 As Double
    Public dpgas2 As Double         '[Pa] pressure loss gas
    Public dpdust2 As Double        '[Pa] pressure loss dust
    Public m2 As Double             'm factor loss curve d< dia critical
    Public stage2() As GvG_Calc_struct   'tbv calculatie stage #2
    Public Dmin2 As Double          'Smallest particle in calculation
    Public Dmax2 As Double          'Biggest particle in calculation
End Structure

'Variables used by GvG in calculation
<Serializable()> Public Structure GvG_Calc_struct
    Public dia As Double            'Particle diameter [mu]
    Public d_ave As Double          'Average diameter [mu]
    Public d_ave_K As Double        'Average diam/K_stokes [-]
    Public loss_overall As Double   'Overall Corrected
    Public loss_overall_C As Double 'Overall loss Corrected
    Public catch_chart As Double    '[%] for chart
    Public i_grp As Double          'Particle Groepnummer
    Public i_d1 As Double           'Class diameter lower[mu]
    Public i_d2 As Double           'Class diameter upper[mu]
    Public i_p1 As Double           'Interpolatie 1
    Public i_p2 As Double           'Interpolatie 2
    Public i_k As Double            'Parameter k
    Public i_m As Double            'Parameter m
    Public psd_cum As Double        '[-] Partice Size Distribution cummulatief
    Public psd_cum_pro As Double    '[%] PSD cummulatief
    Public psd_dif As Double        '[%] PSD diff
    Public loss_abs As Double       '[&] loss abs
    Public loss_abs_C As Double     '[&] loss abs compensated

    Public Shared Widening Operator CType(v As Integer) As GvG_Calc_struct
        Throw New NotImplementedException()
    End Operator
End Structure

Public Class Form1
    Public _cyl1_dim(20) As Double          'Cyclone stage #1 dimensions
    Public _cyl2_dim(20) As Double          'Cyclone stage #2 dimensions
    Public _cees(20) As Input_struct        '20 Case's data


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
    "AC-850+afz;0.203;0.457;0.6;0.564;0.3;0.307;0.428;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-1850;0.136;0.31;0.6;0.53;0.3;0.15;0.25;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-1850+afz;0.136;0.31;0.6;0.53;0.3;0.15;0.25;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25"}

    'Nieuwe reken methode, verdeling volgens Weibull verdeling
    'm1,k1,a1 als d < d_krit
    'm2,k2,a2 als d > d_krit
    'type; d/krit; m1; k1; a1; m2; k2; a2; drukcoef air;drukcoef dust
    Public Shared rekenlijnen() As String = {
    "AC300;     12.2;   1.15;   7.457;  1.005;      8.5308;     1.6102; 0.4789; 7;      0",
    "AC350;     10.2;   1.0;    5.3515; 1.0474;     4.4862;     2.4257; 0.6472; 7;      7.927",
    "AC435;     8.93;   0.69;   4.344;  1.139;      4.2902;     1.3452; 0.5890; 7;      8.26",
    "AC550;     8.62;   0.527;  3.4708; 0.9163;     3.3211;     1.7857; 0.7104; 7;      7.615",
    "AC750;     8.3;    0.50;   2.8803; 0.8355;     4.0940;     1.0519; 0.6010; 7.5;    6.606",
    "AC850;     7.8;    0.52;   1.9418; 0.73705;    -0.1060;    2.0197; 0.7077; 9.5;    6.172",
    "AC850+afz; 10;     0.5187; 1.6412; 0.8386;     4.2781;     0.06777;0.3315; 0;      0",
    "AC1850;    9.3;    0.50;   1.1927; 0.5983;     -0.196;     1.3687; 0.6173; 14.5;   0",
    "AC1850+afz;10.45;  0.4617; 0.2921; 0.4560;     -0.2396;    0.1269; 0.3633; 0;      0"}

    '----------- directory's-----------
    Dim dirpath_Eng As String = "N:\Engineering\VBasic\Cyclone_sizing_cees\"
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

        'Initialize the arrays in the struct
        For i = 0 To _cees.Length - 1
            ReDim _cees(i).dia_big(11)          'Initialize
            ReDim _cees(i).class_load(11)       'Initialize
            ReDim _cees(i).stage1(150)          'Initialize
            ReDim _cees(i).stage2(150)          'Initialize
        Next

        '------ allowed users with hard disc id's -----
        user_list.Add("user")
        hard_disk_list.Add("058F63646471")      'Privee PC, graslaan25

        user_list.Add("GerritP")
        hard_disk_list.Add("S2R6NX0H740154H")       'VTK PC, GP

        user_list.Add("GerritP")
        hard_disk_list.Add("0008_0D02_003E_0FBB.")  'VTK laptop, GP

        user_list.Add("GP")
        hard_disk_list.Add("S28ZNXAG521979")  'VTK laptop, GP privee

        user_list.Add("FredKo")
        hard_disk_list.Add("JR10006P02Y6EE")        'VTK laptop, FKo

        user_list.Add("VittorioS")
        hard_disk_list.Add("002427108605")          'VTK laptop, Vittorio

        user_list.Add("keess")
        hard_disk_list.Add("002410146654")          'VTK laptop, KeesS

        user_list.Add("JanK")
        hard_disk_list.Add("0025_38B4_71B4_88FC.")  'VTK laptop, Jank

        user_list.Add("JeroenA")
        hard_disk_list.Add("171095402070")          'VTK desktop, Jeroen

        user_list.Add("JeroenA")
        hard_disk_list.Add("170228801578")          'VTK laptop, Jeroen disk 1
        hard_disk_list.Add("MCDBM1M4F3QRBEH6")      'VTK laptop, Jeroen disk 2
        hard_disk_list.Add("0025_388A_81BB_14B5.")  'Zweet kamer, Jeroen 

        user_list.Add("lennardh")
        hard_disk_list.Add("141190402709")          'VTK PC, Lennard Hubert

        user_list.Add("Peterdw")
        hard_disk_list.Add("134309552747")          'VTK PC, Peter de Wild

        user_list.Add("Jeffreyvdz")
        hard_disk_list.Add("ACE4_2E81_7006_2BD9.")  'VTK Laptop, Jeffrey van der Zwart

        user_list.Add("Twana")
        hard_disk_list.Add("ACE4_2E81_7006_2BD7.")  'VTK Laptop, Twan Akbheis

        user_list.Add("robru")
        hard_disk_list.Add("174741803447")          'VTK Laptop, Rob Ruiter

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

        For hh = 0 To (cyl_dimensions.Length - 1)  'Fill combobox1 cyclone types
            words = cyl_dimensions(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
            ComboBox2.Items.Add(words(0))
        Next hh
        ComboBox1.SelectedIndex = 2                 'Select Cyclone type AC_435
        ComboBox2.SelectedIndex = 5                 'Select Cyclone type AC_850

        TextBox20.Text = "All AA cyclones have a diameter of 300mm" & vbCrLf
        TextBox20.Text &= "Load above 5 gr/m3 is considered a high load" & vbCrLf
        TextBox20.Text &= "Cyclones can not choke" & vbCrLf

        TextBox47.Text = "Applications" & vbCrLf
        TextBox47.Text &= "Fly catcher Venezuela before a gasturbine" & vbCrLf
        TextBox47.Text &= "Spark catcher (metalic particles)" & vbCrLf
        TextBox47.Text &= "Droplet catcher" & vbCrLf
        TextBox47.Text &= "Patato Starch" & vbCrLf

        Calc_sequence()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles button1.Click, TabPage1.Enter, numericUpDown3.ValueChanged, numericUpDown2.ValueChanged, numericUpDown14.ValueChanged, NumericUpDown1.ValueChanged, numericUpDown5.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown18.ValueChanged, ComboBox1.SelectedIndexChanged, numericUpDown9.ValueChanged, numericUpDown8.ValueChanged, numericUpDown7.ValueChanged, numericUpDown6.ValueChanged, numericUpDown12.ValueChanged, numericUpDown11.ValueChanged, numericUpDown10.ValueChanged, numericUpDown13.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown27.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown34.ValueChanged, NumericUpDown33.ValueChanged, ComboBox2.SelectedIndexChanged, NumericUpDown40.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown38.ValueChanged, NumericUpDown37.ValueChanged, NumericUpDown36.ValueChanged, NumericUpDown35.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown22.ValueChanged, CheckBox4.CheckedChanged
        Calc_sequence()
    End Sub
    Private Sub Get_input_calc_1(ks As Integer)
        Dim db1 As Double           'Body diameter stage #1
        Dim db2 As Double           'Body diameter stage #2
        Dim words() As String

        '===Input parameter ===
        Dim tot_kgh As Double       'Dust inlet per hour totaal 
        Dim ro_gas As Double        'Density [kg/hr]
        Dim ro_solid As Double      'Density [kg/hr]
        Dim visco As Double         'Visco in Centi Poise

        '==== data ======
        Dim wc_dust1, wc_dust2 As Double    'weerstand_coef_air
        Dim wc_air1, wc_air2 As Double      'weerstand_coef_air

        '==== results ===
        Dim kgh As Double           'Dust inlet per hour/cycloon 
        Dim kgs As Double           'Dust inlet per second

        ''==== stage 1 ====
        'Dim h18, h19 As Double
        'Dim j18, i18 As Double
        'Dim l18, k19, k41 As Double
        'Dim k18 As Double
        'Dim m18, n17_oud, n18 As Double
        'Dim tot_catch_abs As Double
        'Dim o18 As Double

        If (ComboBox1.SelectedIndex > -1) And (ComboBox2.SelectedIndex > -1) Then 'Prevent exceptions
            '-------- dimension cyclone stage #1
            words = cyl_dimensions(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            For hh = 1 To 15
                _cyl1_dim(hh) = CDbl(words(hh))          'Cyclone dimensions
            Next

            '-------- dimension cyclone stage #2
            words = cyl_dimensions(ComboBox2.SelectedIndex).Split(CType(";", Char()))
            For hh = 1 To 15
                _cyl2_dim(hh) = CDbl(words(hh))          'Cyclone dimensions
            Next

            _cees(ks).Noc1 = CInt(NumericUpDown20.Value) 'Paralelle cyclonen
            _cees(ks).Noc2 = CInt(NumericUpDown33.Value) 'Paralelle cyclonen
            _cees(ks).Druk1 = NumericUpDown19.Value + 1013.25  'Pressure [mbar]

            db1 = numericUpDown5.Value / 1000            '[m] Body diameter
            db2 = NumericUpDown34.Value / 1000           '[m] Body diameter
            _cees(ks).db1 = db1                          '[m] Body diameter
            _cees(ks).db2 = db2                          '[m] Body diameter
            _cees(ks).inh1 = _cyl1_dim(1) * db1          '[m] inlet hoog
            _cees(ks).inh2 = _cyl2_dim(1) * db2          '[m] inlet hoog
            _cees(ks).inb1 = _cyl1_dim(2) * db1          '[m] inlet breed
            _cees(ks).inb2 = _cyl2_dim(2) * db2          '[m] inlet breed
            _cees(ks).dout1 = _cyl1_dim(6) * db1         '[m] diameter gas outlet
            _cees(ks).dout2 = _cyl2_dim(6) * db2         '[m] diameter gas outlet

            _cees(ks).stofb1 = NumericUpDown4.Value     '[g/Am3]
            CheckBox2.Checked = CBool(IIf(_cees(ks).stofb1 > 5, True, False))
            _cees(ks).FlowT = NumericUpDown1.Value      '[m3/h] 
            _cees(ks).Flow1 = _cees(ks).FlowT / (3600 * _cees(ks).Noc1) '[Am3/s/cycloon]

            ro_gas = numericUpDown3.Value               '[kg/m3]
            ro_solid = numericUpDown2.Value             '[kg/m3]
            visco = numericUpDown14.Value               '[cPoise]

            '=========== Stage #1 ==============
            _cees(ks).inv1 = _cees(ks).Flow1 / (_cees(ks).inb1 * _cees(ks).inh1)
            _cees(ks).outv1 = _cees(ks).Flow1 / ((PI / 4) * _cees(ks).dout1 ^ 2)   '[m/s]

            If ComboBox1.SelectedIndex > -1 Then
                words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
                wc_air1 = CDbl(words(8))        'Resistance Coefficient air
                wc_dust1 = CDbl(words(9))       'Resistance Coefficient dust
            End If
            _cees(ks).dpgas1 = 0.5 * ro_gas * _cees(ks).inv1 ^ 2 * wc_air1
            _cees(ks).dpdust1 = 0.5 * ro_gas * _cees(ks).inv1 ^ 2 * wc_dust1

            '=========== Stage #2 ==============
            _cees(ks).Druk2 = _cees(ks).Druk1 - _cees(ks).dpgas1 / 100  '[mbar]
            _cees(ks).Flow2 = _cees(ks).FlowT / (3600 * _cees(ks).Noc2) '[Am3/s/cycloon]
            If Not CheckBox4.Checked Then   'Stage 1 is Bypassed
                _cees(ks).Flow2 *= _cees(ks).Druk1 / _cees(ks).Druk2    '[Am3/s/cycloon]
            End If


            '---- Compensate for the Flow for the pressure loss in stage #1 ----
            _cees(ks).inv2 = _cees(ks).Flow2 / (_cees(ks).inb2 * _cees(ks).inh2)
            _cees(ks).outv2 = _cees(ks).Flow2 / ((PI / 4) * _cees(ks).dout2 ^ 2)   '[m/s]

            'TextBox24.Text &= "_cees(ks).Flow2= " & _cees(ks).Flow2.ToString
            'TextBox24.Text &= ",  _cees(ks).inb2=" & _cees(ks).inb2.ToString
            'TextBox24.Text &= ",  _cees(ks).inh2=" & _cees(ks).inh2.ToString
            'TextBox24.Text &= ",   _cees(ks).inv2 =" & _cees(ks).inv2.ToString & vbCrLf

            '----------- Pressure loss cyclone stage #2----------------------
            If ComboBox2.SelectedIndex > -1 Then
                words = rekenlijnen(ComboBox2.SelectedIndex).Split(CType(";", Char()))
                wc_air2 = CDbl(words(8))
                wc_dust2 = CDbl(words(9))
            End If
            _cees(ks).dpgas2 = 0.5 * ro_gas * _cees(ks).inv2 ^ 2 * wc_air2
            _cees(ks).dpdust2 = 0.5 * ro_gas * _cees(ks).inv2 ^ 2 * wc_dust2

            '----------- stof belasting ------------
            kgs = _cees(ks).Flow1 * _cees(ks).stofb1 / 1000     '[kg/s/cycloon]
            kgh = kgs * 3600                                    '[kg/h/cycloon]
            tot_kgh = kgh * _cees(ks).Noc1                      '[g/Am3] Dust inlet 

            '----------- K_stokes-----------------------------------
            _cees(ks).Kstokes1 = Sqrt(db1 * 2000 * visco * 16 / (ro_solid * 0.0181 * _cees(ks).inv1))
            _cees(ks).Kstokes2 = Sqrt(db2 * 2000 * visco * 16 / (ro_solid * 0.0181 * _cees(ks).inv2))

            TextBox24.Text &= "db2= " & db2.ToString
            TextBox24.Text &= ",  visco=" & visco.ToString
            TextBox24.Text &= ",  ro_solid=" & ro_solid.ToString
            TextBox24.Text &= ",  _cees(ks).inv2=" & _cees(ks).inv2.ToString
            TextBox24.Text &= ",  _cees(ks).Kstokes2=" & _cees(ks).Kstokes2.ToString & vbCrLf


            '----------- presenteren ----------------------------------
            TextBox36.Text = (_cees(ks).FlowT / 3600).ToString("F4")    '[m3/s] flow

            '----------- presenteren afmetingen ------------------------------
            TextBox1.Text = (_cees(ks).inh1).ToString("0.000")          'inlaat breedte
            TextBox2.Text = (_cees(ks).inb1).ToString("0.000")          'Inlaat hoogte
            TextBox3.Text = (_cyl1_dim(3) * db1).ToString("0.000")      'Inlaat lengte
            TextBox4.Text = (_cyl1_dim(4) * db1).ToString("0.000")      'Inlaat hartmaat
            TextBox5.Text = (_cyl1_dim(5) * db1).ToString("0.000")      'Inlaat afschuining
            TextBox6.Text = (_cyl1_dim(6) * db1).ToString("0.000")      'Uitlaat keeldia inw.
            TextBox7.Text = (_cyl1_dim(7) * db1).ToString("0.000")      'Uitlaat flensdiameter inw.
            TextBox8.Text = (_cyl1_dim(8) * db1).ToString("0.000")      'Lengte insteekpijp inw.
            TextBox9.Text = (_cyl1_dim(9) * db1).ToString("0.000")      'Lengte romp + conus
            TextBox10.Text = (_cyl1_dim(10) * db1).ToString("0.000")    'Lengte romp
            TextBox11.Text = (_cyl1_dim(11) * db1).ToString("0.000")    'Lengte çonus
            TextBox12.Text = (_cyl1_dim(12) * db1).ToString("0.000")    'Dia_conus / 3P-pijp
            TextBox13.Text = (_cyl1_dim(13) * db1).ToString("0.000")    'Lengte 3P-pijp
            TextBox14.Text = (_cyl1_dim(14) * db1).ToString("0.000")    'Lengte 3P conus
            TextBox15.Text = (_cyl1_dim(15) * db1).ToString("0.000")    'Kleine dia 3P-conus

            TextBox84.Text = (_cees(ks).inh2).ToString("0.000")         'inlaat breedte
            TextBox85.Text = (_cees(ks).inb2).ToString("0.000")         'Inlaat hoogte
            TextBox86.Text = (_cyl1_dim(3) * db2).ToString("0.000")     'Inlaat lengte
            TextBox87.Text = (_cyl1_dim(4) * db2).ToString("0.000")     'Inlaat hartmaat
            TextBox88.Text = (_cyl1_dim(5) * db2).ToString("0.000")     'Inlaat afschuining
            TextBox89.Text = (_cyl1_dim(6) * db2).ToString("0.000")     'Uitlaat keeldia inw.
            TextBox90.Text = (_cyl1_dim(7) * db2).ToString("0.000")     'Uitlaat flensdiameter inw.
            TextBox91.Text = (_cyl1_dim(8) * db2).ToString("0.000")     'Lengte insteekpijp inw.
            TextBox92.Text = (_cyl1_dim(9) * db2).ToString("0.000")     'Lengte romp + conus
            TextBox93.Text = (_cyl1_dim(10) * db2).ToString("0.000")    'Lengte romp
            TextBox94.Text = (_cyl1_dim(11) * db2).ToString("0.000")    'Lengte çonus
            TextBox95.Text = (_cyl1_dim(12) * db2).ToString("0.000")    'Dia_conus / 3P-pijp
            TextBox96.Text = (_cyl1_dim(13) * db2).ToString("0.000")    'Lengte 3P-pijp
            TextBox97.Text = (_cyl1_dim(14) * db2).ToString("0.000")    'Lengte 3P conus
            TextBox98.Text = (_cyl1_dim(15) * db2).ToString("0.000")    'Kleine dia 3P-conus

            TextBox113.Text = (_cees(ks).Flow1 * 3600).ToString("0")    '[Am3/s] Cycloone Flow
            TextBox112.Text = (_cees(ks).Flow2 * 3600).ToString("0")    '[Am3/s] Cycloone Flow

            TextBox16.Text = _cees(ks).inv1.ToString("0.0")             'inlaat snelheid
            TextBox80.Text = _cees(ks).inv2.ToString("0.0")             'inlaat snelheid
            TextBox17.Text = _cees(ks).dpgas1.ToString("0")             '[Pa] Pressure loss inlet-gas
            TextBox19.Text = (_cees(ks).dpgas1 / 100).ToString("0.0")   '[mbar] Pressure loss inlet-gas
            TextBox48.Text = _cees(ks).dpdust1.ToString("0")            '[Pa]Pressure loss inlet-dust

            TextBox79.Text = _cees(ks).dpgas2.ToString("0")             '[Pa] Pressure loss inlet-gas
            TextBox75.Text = (_cees(ks).dpgas2 / 100).ToString("0.0")   '[mbar] Pressure loss inlet-gas
            TextBox76.Text = _cees(ks).dpdust2.ToString("0")            '[Pa]Pressure loss inlet-dust

            TextBox22.Text = _cees(ks).outv1.ToString("0.0")            'uitlaat snelheid
            TextBox77.Text = _cees(ks).outv2.ToString("0.0")            'uitlaat snelheid

            TextBox23.Text = _cees(ks).Kstokes1.ToString("F4")       'Stokes waarde stage#1
            TextBox78.Text = _cees(ks).Kstokes2.ToString("F4")       'Stokes waarde stage#1

            TextBox37.Text = _cees(ks).db1.ToString                     'Cycloone diameter
            TextBox74.Text = _cees(ks).db2.ToString                     'Cycloone diameter

            TextBox38.Text = CType(ComboBox1.SelectedItem, String)      'Cycloon type
            TextBox73.Text = CType(ComboBox2.SelectedItem, String)      'Cycloon type

            Draw_chart1(Chart1)
            Draw_chart2(Chart2)
            '---------- Check speed stage #1---------------
            If _cees(ks).inv1 < 10 Or _cees(ks).inv1 > 30 Then
                TextBox16.BackColor = Color.Red
            Else
                TextBox16.BackColor = Color.LightGreen
            End If

            '---------- Check speed stage #2---------------
            If _cees(ks).inv2 < 10 Or _cees(ks).inv2 > 30 Then
                TextBox80.BackColor = Color.Red
            Else
                TextBox80.BackColor = Color.LightGreen
            End If

            '---------- Check dp stage #1---------------
            If _cees(ks).dpgas1 > 3000 Then
                TextBox17.BackColor = Color.Red
                TextBox19.BackColor = Color.Red
            Else
                TextBox17.BackColor = Color.LightGreen
                TextBox19.BackColor = Color.LightGreen
            End If

            '---------- Check dp stage #2---------------
            If _cees(ks).dpgas2 > 3000 Then
                TextBox79.BackColor = Color.Red
                TextBox75.BackColor = Color.Red
            Else
                TextBox79.BackColor = Color.LightGreen
                TextBox75.BackColor = Color.LightGreen
            End If



            '--------- Get Inlet korrel-groep data ----------
            'Save data of screen into the _cees array
            Fill_cees_array(CInt(NumericUpDown30.Value))


            TextBox39.Text = kgh.ToString("0")          'Stof inlet
            TextBox40.Text = tot_kgh.ToString("0")      'Dust inlet [g/Am3] 
            TextBox71.Text = _cees(ks).stofb1.ToString  'Dust inlet [g/Am3]

        End If
    End Sub
    Private Sub Present_Datagridview1(ks As Integer)
        '==== stage 1 ====
        Dim h18, h19 As Double
        Dim j18, i18 As Double
        Dim l18, k19, k41 As Double
        Dim k18 As Double
        Dim m18, n17_oud, n18 As Double
        Dim tot_catch_abs As Double
        Dim o18 As Double

        '--------- overall resultaat --------------------
        DataGridView1.Columns(0).HeaderText = "Dia class"
        DataGridView1.Columns(1).HeaderText = "Feed psd cum"
        DataGridView1.Columns(2).HeaderText = "Feed psd diff"
        DataGridView1.Columns(3).HeaderText = "Loss % of feed"
        DataGridView1.Columns(4).HeaderText = "Loss abs [%]"
        DataGridView1.Columns(5).HeaderText = "Loss psd cum"
        DataGridView1.Columns(6).HeaderText = "Catch abs"
        DataGridView1.Columns(7).HeaderText = "Catch psd cum"
        DataGridView1.Columns(8).HeaderText = "Grade class eff."

        For h = 0 To 22
            DataGridView1.Rows.Item(h).Cells(0).Value = _cees(ks).stage1(h * 5).d_ave.ToString("0.000") 'diameter
            DataGridView1.Rows.Item(h).Cells(1).Value = _cees(ks).stage1(h * 5).psd_cum_pro.ToString("0.0") 'feed psd cum

            If h > 0 Then
                h18 = CDbl(DataGridView1.Rows.Item(h - 1).Cells(1).Value)
            Else
                h18 = 100
            End If
            h19 = CDbl(DataGridView1.Rows.Item(h).Cells(1).Value)   'feed psd cum
            DataGridView1.Rows.Item(h).Cells(2).Value = (h18 - h19).ToString("0.00")   'feed psd diff

            '========= loss ===============
            If CheckBox1.Checked Then
                DataGridView1.Rows.Item(h).Cells(3).Value = (_cees(ks).stage1(h * 5).loss_overall * 100).ToString("0.0000")
            Else
                DataGridView1.Rows.Item(h).Cells(3).Value = (_cees(ks).stage1(h * 5).loss_overall_C * 100).ToString("0.0000")
            End If

            i18 = CDbl(DataGridView1.Rows.Item(h).Cells(2).Value) 'feed psd diff
            j18 = CDbl(DataGridView1.Rows.Item(h).Cells(3).Value) 'loss % Of feed
            DataGridView1.Rows.Item(h).Cells(4).Value = (i18 * j18 / 100).ToString("0.0000") 'Loss abs [%]
            If h > 0 Then
                l18 = CDbl(DataGridView1.Rows.Item(h - 1).Cells(5).Value)
            Else
                l18 = 100
            End If
            k19 = CDbl(DataGridView1.Rows.Item(h).Cells(4).Value)   'Loss abs [%]

            '========= Catch ===============
            Double.TryParse(TextBox58.Text, k41)
            DataGridView1.Rows.Item(h).Cells(5).Value = (l18 - 100 * k19 / k41).ToString("0.0000")

            If h > 0 Then
                l18 = CDbl(DataGridView1.Rows.Item(h - 1).Cells(5).Value)
            Else
                l18 = 100
            End If

            k18 = CDbl(DataGridView1.Rows.Item(h).Cells(4).Value)   'Loss abs [%]
            m18 = (i18 - k18)
            DataGridView1.Rows.Item(h).Cells(6).Value = m18.ToString("0.000") 'Catch abs

            Double.TryParse(TextBox59.Text, tot_catch_abs)      'tot_catch_abs[%]

            If h > 0 Then
                n17_oud = CDbl(DataGridView1.Rows.Item(h - 1).Cells(7).Value)
                n18 = n17_oud - m18 / (tot_catch_abs / 100)
            Else
                n18 = 100
            End If

            n18 = CDbl(IIf(n18 < 0, 0, n18))        'prevent silly results
            'TextBox24.Text &= "**h= " & h.ToString & ", n17_oud= " & n17_oud.ToString
            'TextBox24.Text &= ", m18= " & m18.ToString & ",==> n18= " & n18.ToString & vbCrLf
            DataGridView1.Rows.Item(h).Cells(7).Value = n18.ToString("0.000") 'Catch psd cum

            '========= Efficiency ===============
            o18 = 100 - j18
            DataGridView1.Rows.Item(h).Cells(8).Value = o18.ToString("0.000")           'Grade eff.
        Next h
        DataGridView1.AutoResizeColumns()
    End Sub

    Private Sub Calc_part_dia_loss(ks As Integer)

        '---------- Calc particle diameter with x% loss ---
        '---------- present stage #1 -------
        TextBox42.Text = Calc_dia_particle(1.0, _cees(ks).Kstokes1, 1).ToString("0.00")     '[mu] @ 100% loss
        TextBox26.Text = Calc_dia_particle(0.95, _cees(ks).Kstokes1, 1).ToString("0.00")    '[mu] @  95% lost
        TextBox31.Text = Calc_dia_particle(0.9, _cees(ks).Kstokes1, 1).ToString("0.00")     '[mu] @  90% lost
        TextBox32.Text = Calc_dia_particle(0.5, _cees(ks).Kstokes1, 1).ToString("0.00")     '[mu] @  50% lost
        TextBox33.Text = Calc_dia_particle(0.1, _cees(ks).Kstokes1, 1).ToString("0.00")     '[mu] @  10% lost
        TextBox41.Text = Calc_dia_particle(0.05, _cees(ks).Kstokes1, 1).ToString("0.00")    '[mu] @   5% lost

        '---------- present stage #2 -------
        ' MessageBox.Show(_cees(ks).Kstokes2.ToString)
        TextBox102.Text = Calc_dia_particle(1.0, _cees(ks).Kstokes2, 2).ToString("0.00")     '[mu] @  100% lost
        TextBox103.Text = Calc_dia_particle(0.95, _cees(ks).Kstokes2, 2).ToString("0.00")    '[mu] @  95% lost
        TextBox104.Text = Calc_dia_particle(0.9, _cees(ks).Kstokes2, 2).ToString("0.00")     '[mu] @  90% lost
        TextBox105.Text = Calc_dia_particle(0.5, _cees(ks).Kstokes2, 2).ToString("0.00")     '[mu] @  50% lost
        TextBox106.Text = Calc_dia_particle(0.1, _cees(ks).Kstokes2, 2).ToString("0.00")     '[mu] @  10% lost
        TextBox107.Text = Calc_dia_particle(0.05, _cees(ks).Kstokes2, 2).ToString("0.00")    '[mu] @   5% lost
    End Sub

    Private Sub Fill_cees_array(c_nr As Integer)
        DataGridView1.ColumnCount = 10
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(23)
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

        _cees(c_nr).case_name = TextBox53.Text          'The case name

        '[mu] Class upper particle diameter limit diameter
        _cees(c_nr).dia_big(0) = NumericUpDown15.Value   '10
        _cees(c_nr).dia_big(1) = NumericUpDown23.Value   '15
        _cees(c_nr).dia_big(2) = NumericUpDown24.Value   '20
        _cees(c_nr).dia_big(3) = NumericUpDown25.Value   '30
        _cees(c_nr).dia_big(4) = NumericUpDown26.Value   '40
        _cees(c_nr).dia_big(5) = NumericUpDown27.Value   '50
        _cees(c_nr).dia_big(6) = NumericUpDown28.Value   '60
        _cees(c_nr).dia_big(7) = NumericUpDown29.Value   '80
        _cees(c_nr).dia_big(8) = NumericUpDown35.Value   '50
        _cees(c_nr).dia_big(9) = NumericUpDown36.Value   '60
        _cees(c_nr).dia_big(10) = NumericUpDown37.Value   '80


        'Percentale van de inlaat stof belasting
        _cees(c_nr).class_load(0) = numericUpDown6.Value / 100
        _cees(c_nr).class_load(1) = numericUpDown7.Value / 100
        _cees(c_nr).class_load(2) = numericUpDown8.Value / 100
        _cees(c_nr).class_load(3) = numericUpDown9.Value / 100
        _cees(c_nr).class_load(4) = numericUpDown10.Value / 100
        _cees(c_nr).class_load(5) = numericUpDown11.Value / 100
        _cees(c_nr).class_load(6) = numericUpDown12.Value / 100
        _cees(c_nr).class_load(7) = numericUpDown13.Value / 100
        _cees(c_nr).class_load(8) = NumericUpDown38.Value / 100
        _cees(c_nr).class_load(9) = NumericUpDown39.Value / 100
        _cees(c_nr).class_load(10) = NumericUpDown40.Value / 100

        _cees(c_nr).FlowT = NumericUpDown1.Value            'Air flow
        _cees(c_nr).stofb1 = NumericUpDown4.Value           'Dust inlet [g/Am3] 
        _cees(c_nr).Ct1 = ComboBox1.SelectedIndex           'Cyclone type Stage #1
        _cees(c_nr).Ct2 = ComboBox2.SelectedIndex           'Cyclone type Stage #2
        _cees(c_nr).Noc1 = CInt(NumericUpDown20.Value)      'Cyclone in parallel
        _cees(c_nr).Noc2 = CInt(NumericUpDown33.Value)      'Cyclone in parallel
        _cees(c_nr).db1 = numericUpDown13.Value             'Diameter cyclone Stage #1
        _cees(c_nr).db2 = NumericUpDown34.Value             'Diameter cyclone Stage #2
        _cees(c_nr).ro_gas = numericUpDown3.Value           'Density [kg/hr]
        _cees(c_nr).ro_solid = numericUpDown2.Value         'Density [kg/hr]
        _cees(c_nr).visco = numericUpDown14.Value           'Visco in Centi Poise
        _cees(c_nr).Temp = NumericUpDown18.Value            'Temperature [c]
        _cees(c_nr).Druk1 = NumericUpDown19.Value           'Pressure [mbar]

        '-------- Check -- bigger diameter must have bigger cummulative weight
        numericUpDown6.BackColor = CType(IIf(numericUpDown6.Value > numericUpDown7.Value, Color.LightGreen, Color.Red), Color)
        numericUpDown7.BackColor = CType(IIf(numericUpDown7.Value > numericUpDown8.Value, Color.LightGreen, Color.Red), Color)
        numericUpDown8.BackColor = CType(IIf(numericUpDown8.Value > numericUpDown9.Value, Color.LightGreen, Color.Red), Color)
        numericUpDown9.BackColor = CType(IIf(numericUpDown9.Value > numericUpDown10.Value, Color.LightGreen, Color.Red), Color)
        numericUpDown10.BackColor = CType(IIf(numericUpDown10.Value > numericUpDown11.Value, Color.LightGreen, Color.Red), Color)
        numericUpDown11.BackColor = CType(IIf(numericUpDown11.Value > numericUpDown12.Value, Color.LightGreen, Color.Red), Color)
        numericUpDown12.BackColor = CType(IIf(numericUpDown12.Value > numericUpDown13.Value, Color.LightGreen, Color.Red), Color)
        numericUpDown13.BackColor = CType(IIf(numericUpDown13.Value > NumericUpDown38.Value, Color.LightGreen, Color.Red), Color)

        NumericUpDown38.BackColor = CType(IIf(NumericUpDown38.Value > NumericUpDown39.Value, Color.LightGreen, Color.Red), Color)
        NumericUpDown39.BackColor = CType(IIf(NumericUpDown39.Value > NumericUpDown40.Value, Color.LightGreen, Color.Red), Color)
        NumericUpDown40.BackColor = CType(IIf(NumericUpDown40.Value >= 0, Color.LightGreen, Color.Red), Color)
    End Sub
    '-------- Bereken het verlies getal NIET gecorrigeerd -----------
    '----- de input is de GEMIDDELDE korrel grootte-----------
    Private Function Calc_verlies(korrel_g As Double, present As Boolean, stokes As Double, stage As Integer) As Double
        Dim words() As String
        Dim dia_Kcrit, fac_m, fac_a, fac_k As Double
        Dim verlies As Double = 1
        Dim dia_K As Double

        If (ComboBox1.SelectedIndex > -1) Then
            '-------------- korrelgrootte factoren ------
            If (stage = 1) Then     'Stage #1 cyclone
                words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
            Else                    'Stage #2 cyclone
                words = rekenlijnen(ComboBox2.SelectedIndex).Split(CType(";", Char()))
            End If

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

            dia_K = korrel_g / stokes
            If (dia_K - fac_m) > 0 Then
                verlies = Math.E ^ -((dia_K - fac_m) / fac_k) ^ fac_a
                'verlies = verlies
            Else
                verlies = 1.0        '100% loss (very small particle)
            End If

            '----- Stage #1 is bypassed ------
            If CheckBox4.Checked And stage = 1 Then
                verlies = 1.0       'Cyclone stage 1 is bypassed ! (100% loss)
            End If
        End If
        Return (verlies)
    End Function

    '-------- Bereken het verlies getal GECORRIGEERD -----------
    '----- de input is de GEMIDDELDE korrel grootte-----------
    Private Sub Calc_verlies_corrected(ByRef grp As GvG_Calc_struct, stage As Integer)
        Dim cor1, cor2 As Double

        If stage > 2 Or stage < 1 Then MessageBox.Show("Problem in Line 658")  '----- check input ----

        If (ComboBox1.SelectedIndex > -1) And ComboBox2.SelectedIndex > -1 Then
            If (stage = 1) Then     'Stage #1 cyclone
                cor1 = NumericUpDown22.Value    'Correctie insteek pijp stage #1
                Double.TryParse(TextBox55.Text, cor2) 'Hoge stof belasting correctie acc VT-UK

            Else                    'Stage #2 cyclone
                cor1 = NumericUpDown43.Value    'Correctie insteek pijp stage #2
                Double.TryParse(TextBox67.Text, cor2) 'Hoge stof belasting correctie acc VT-UK
            End If
            grp.loss_overall_C = grp.loss_overall ^ (cor1 * cor2)
        End If
    End Sub

    'Note dp(95) meaning with this diameter 95% is lost
    'Calculate the diameter at which qq% is lost
    'Separation depends on Stokes
    Private Function Calc_dia_particle(qq As Double, stokes As Double, stage As Integer) As Double
        Dim dia_result As Double = 0
        Dim words() As String
        Dim dia_Kcrit As Double
        Dim d1, d2 As Double
        Dim cor1, cor2 As Double 'Insteek pijp
        Dim fac_m1, fac_k1, fac_a1 As Double
        Dim fac_m2, fac_k2, fac_a2 As Double

        '--- check input ----
        If qq > 1 Then MessageBox.Show("Loss > 100% is impossible, Line 486, qq= " & qq.ToString)
        If stage > 2 Or stage < 1 Then MessageBox.Show("Problem in Line 708")

        '----- Insteek pijp corectie correctie -------
        If (stage = 1) Then     'Cyclone 1# stage
            cor1 = NumericUpDown22.Value    'Correctie insteek pijp stage #1
            Double.TryParse(TextBox55.Text, cor2) 'Hoge stof belasting correctie acc VT-UK
            words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
        Else                    'Cyclone 2# stage
            cor1 = NumericUpDown43.Value    'Correctie insteek pijp stage #2
            Double.TryParse(TextBox67.Text, cor2) 'Hoge stof belasting correctie acc VT-UK
            words = rekenlijnen(ComboBox2.SelectedIndex).Split(CType(";", Char()))
        End If

        '-------------- korrelgrootte factoren ------
        dia_Kcrit = CDbl(words(1))   'Is in fact d/K(crit)

        '---- diameter particle kleiner dan de diameter kritisch
        fac_m1 = CDbl(words(2))
        fac_k1 = CDbl(words(3))
        fac_a1 = CDbl(words(4))
        d1 = fac_k1 * stokes * ((-Math.Log(qq ^ (1 / (cor1 * cor2))))) ^ (1 / fac_a1) + fac_m1 * stokes

        '---- diameter particle groter dan de diameter kritisch
        fac_m2 = CDbl(words(5))
        fac_k2 = CDbl(words(6))
        fac_a2 = CDbl(words(7))
        d2 = fac_k2 * stokes * ((-Math.Log(qq ^ (1 / (cor1 * cor2))))) ^ (1 / fac_a2) + fac_m2 * stokes

        If ((d1 / stokes) < dia_Kcrit) Then
            dia_result = d1     'diameter kleiner kritisch
        Else
            dia_result = d2     'diameter groter kritisch
        End If

        Return (dia_result)
    End Function

    '---- According to VT-UK -----
    Private Sub Dust_load_correction(ks As Integer)
        Dim f1, f2, f3, f4, f, f_used As Double
        Dim dst As Double

        '============ stage 1 cyclone ==========
        dst = _cees(ks).stofb1 / 1000 'Dust load dimension is [kg/Am3]

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

        TextBox55.Text = f_used.ToString("F3")

        '============ stage 2 cyclone ==========
        dst = _cees(ks).stofb2 / 1000 'Dust load dimension is [kg/Am3]
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

        If (CheckBox3.Checked) Then
            f_used = f
        Else
            f_used = 1
        End If

        TextBox67.Text = f_used.ToString("F3")
    End Sub

    Private Sub Draw_chart1(ch As Chart)
        '-------
        Dim s_points(100, 2) As Double
        Dim h As Integer
        Dim sdia As Integer
        Dim ks As Integer

        ch.Series.Clear()
        ch.ChartAreas.Clear()
        ch.Titles.Clear()
        ch.ChartAreas.Add("ChartArea0")

        ch.Series.Add("Series" & h.ToString)
        ch.Series(h).ChartArea = "ChartArea0"
        ch.Series(h).ChartType = DataVisualization.Charting.SeriesChartType.Line
        ch.Series(h).BorderWidth = 2
        ch.Series(h).IsVisibleInLegend = False

        ch.Titles.Add("Loss Curve")
        ch.ChartAreas("ChartArea0").AxisX.Title = "particle dia [mu]"

        ch.ChartAreas("ChartArea0").AxisY.Title = "Loss [%] (niet gevangen)"
        ch.ChartAreas("ChartArea0").AxisY.Minimum = 0       'Loss
        ch.ChartAreas("ChartArea0").AxisY.Maximum = 100     'Loss
        ch.ChartAreas("ChartArea0").AxisY.Interval = 10     'Interval
        ch.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
        ch.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True
        ch.ChartAreas("ChartArea0").AxisX.MinorGrid.Enabled = True
        ch.ChartAreas("ChartArea0").AxisY.MinorGrid.Enabled = True

        If CheckBox1.Checked Then
            ch.ChartAreas("ChartArea0").AxisX.IsLogarithmic = True
            ch.ChartAreas("ChartArea0").AxisX.Minimum = 1     'Particle size
            ch.ChartAreas("ChartArea0").AxisX.Maximum = 100   'Particle size
        Else
            ch.ChartAreas("ChartArea0").AxisX.IsLogarithmic = False
            ch.ChartAreas("ChartArea0").AxisX.Minimum = 0     'Particle size
            ch.ChartAreas("ChartArea0").AxisX.Maximum = 20    'Particle size
        End If

        '----- now calc chart points --------------------------
        Integer.TryParse(TextBox42.Text, sdia)
        ks = CInt(NumericUpDown30.Value)
        s_points(0, 0) = sdia   'Particle diameter [mu]
        s_points(0, 1) = 100    '100% loss
        For h = 1 To 40
            s_points(h, 0) = h                                   'Particle diameter [mu]
            s_points(h, 1) = Calc_verlies(s_points(h, 0), False, _cees(ks).Kstokes1, 1) * 100  'Loss [%]
        Next

        '------ now present-------------
        For h = 0 To 40 - 1   'Fill line chart
            ch.Series(0).Points.AddXY(s_points(h, 0), s_points(h, 1))
        Next h
    End Sub
    Private Sub Draw_chart2(ch As Chart)
        'Small chart on the first tab
        Dim s_points(100, 2) As Double
        Dim h As Integer
        Dim sdia As Integer
        Dim ks As Integer   'Case number

        ch.Series.Clear()
        ch.ChartAreas.Clear()
        ch.Titles.Clear()
        ch.ChartAreas.Add("ChartArea0")

        ch.Series.Add("Series" & h.ToString)
        ch.Series(h).ChartArea = "ChartArea0"
        ch.Series(h).ChartType = DataVisualization.Charting.SeriesChartType.Line
        ch.Series(h).BorderWidth = 2
        ch.Series(h).IsVisibleInLegend = False

        ch.Titles.Add("Loss Curve")
        ch.ChartAreas("ChartArea0").AxisX.Title = "particle dia [mu]"
        ch.ChartAreas("ChartArea0").AxisY.Minimum = 0     'Loss
        ch.ChartAreas("ChartArea0").AxisY.Maximum = 100   'Loss
        ch.ChartAreas("ChartArea0").AxisX.Minimum = 0     'Particle size
        ch.ChartAreas("ChartArea0").AxisX.Maximum = 20    'Particle size

        '----- now calc chart poins --------------------------
        Integer.TryParse(TextBox42.Text, sdia)
        s_points(0, 0) = sdia   'Particle diameter [mu]
        s_points(0, 1) = 100    '100% loss
        ks = CInt(NumericUpDown30.Value)
        For h = 1 To 40
            s_points(h, 0) = h                                   'Particle diameter [mu]
            s_points(h, 1) = Calc_verlies(s_points(h, 0), False, _cees(ks).Kstokes2, 2) * 100  'Loss [%]
        Next

        '------ now present-------------
        For h = 0 To 40 - 1   'Fill line chart
            ch.Series(0).Points.AddXY(s_points(h, 0), s_points(h, 1))
        Next h
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click, TabPage9.Enter, CheckBox1.CheckedChanged
        Calc_sequence()
    End Sub
    Private Sub Calc_sequence()
        Dim case_nr As Integer = CInt(NumericUpDown30.Value)

        If ComboBox1.SelectedIndex > -1 And ComboBox2.SelectedIndex > -1 Then
            Dust_load_correction(case_nr)
            Get_input_calc_1(case_nr)   'This is the CASE number
            Calc_part_dia_loss(case_nr)
            Calc_stage1(case_nr)        'Calc according stage #1
            Calc_stage2(case_nr)        'Calc according stage #2
            Present_loss_grid1()        'Present the results stage #1
            Present_loss_grid2()        'Present the results stage #2
            Draw_chart1(Chart1)         'Present the results
            Present_Datagridview1(case_nr)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox28.Text.Trim.Length > 0 And TextBox29.Text.Trim.Length > 0 Then
            'Save_tofile()
            Save_tofile2()
        Else
            MessageBox.Show("Complete Quote and Tag number")
        End If
    End Sub
    Private Sub Save_tofile2()
        Dim filename, user As String
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        '------------- create filemame ----------
        user = Trim(Environment.UserName)         'User name on the screen
        filename = "Cyclone_select_" & TextBox28.Text & "_" & TextBox29.Text & DateTime.Now.ToString("_yyyy_MM_dd_") & user & ".vtk2"
        filename = Replace(filename, Chr(32), Chr(95)) 'Replace the space's

        MessageBox.Show(filename)

        Try
            If Directory.Exists(dirpath_Eng) Then
                filename = dirpath_Eng & filename 'used at VTK with intranet
            Else
                filename = dirpath_tmp & filename 'used at VTK with intranet'used at home
            End If
            Dim fStream As New FileStream(filename, FileMode.OpenOrCreate)
            bf.Serialize(fStream, _cees) ' write to file
        Catch ex As Exception
            MessageBox.Show("Line 6298, " & ex.Message)  ' Show the exception's message.
        End Try

        ' fStream.Position = 0 ' reset stream pointer
        ' _cees = CType(bf.Deserialize(fStream), Input_struct()) ' read from file
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

        '-------- Case information -----------------
        For j = 0 To _cees.GetLength(0) - 2                     '20 elements
            temp_string &= _cees(j).FlowT.ToString & ";"        'Air flow
            temp_string &= _cees(j).stofb1.ToString & ";"       'Dust inlet [g/Am3] 
            temp_string &= _cees(j).Ct1.ToString & ";"          'Cyclone type
            temp_string &= _cees(j).Noc1.ToString & ";"         'Cyclone in parallel
            temp_string &= _cees(j).db1.ToString & ";"          'Diameter cyclone body
            temp_string &= _cees(j).ro_gas.ToString & ";"       'Density [kg/hr]
            temp_string &= _cees(j).ro_solid.ToString & ";"     'Density [kg/hr]
            temp_string &= _cees(j).visco.ToString & ";"        'Visco in Centi Poise
            temp_string &= _cees(j).Temp.ToString & ";"         'Temperature [c]
            temp_string &= _cees(j).Druk1.ToString & ";"        'Pressure [mbar]

            For k = 0 To 7         '8 elements
                temp_string &= _cees(j).dia_big(k).ToString & ";"       'Write all variables
            Next

            For k = 0 To 7         '8 elements
                temp_string &= _cees(j).class_load(k).ToString & ";"   'Write all variables
            Next
        Next
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
        Dim count As Integer

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

            count = 2
            Try
                '-------- Case information -----------------
                For j = 0 To _cees.GetLength(0) - 2

                    _cees(j).FlowT = CDbl(words(count))     'Air flow total
                    count += 1
                    _cees(j).stofb1 = CDbl(words(count))    'Dust inlet [g/Am3] 
                    count += 1
                    _cees(j).Ct1 = CInt(words(count))       'Cyclone type
                    count += 1
                    _cees(j).Noc1 = CInt(words(count))    'Cyclone in parallel
                    count += 1
                    _cees(j).db1 = CDbl(words(count))       'Diameter cyclone body
                    count += 1
                    _cees(j).ro_gas = CDbl(words(count))   'Density [kg/hr]
                    count += 1
                    _cees(j).ro_solid = CDbl(words(count)) 'Density [kg/hr]
                    count += 1
                    _cees(j).visco = CDbl(words(count))    'Visco in Centi Poise
                    count += 1
                    _cees(j).Temp = CDbl(words(count))     'Temperature [c]
                    count += 1
                    _cees(j).Druk1 = CDbl(words(count))    'Pressure [mbar]
                    count += 1

                    For k = 0 To 7         '8 elements
                        _cees(j).dia_big(k) = CDbl(words(count))    'Write all variables
                        count += 1
                    Next

                    For k = 0 To 7         '8 elements
                        _cees(j).class_load(k) = CDbl(words(count))     'Write all variables
                        count += 1
                    Next
                Next
            Catch ex As Exception
                MessageBox.Show("Line 923, " & ex.Message)
            End Try

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
            Calc_sequence()
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
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 16, 3)
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
            oTable.Cell(row, 1).Range.Text = "Emission"
            oTable.Cell(row, 2).Range.Text = TextBox18.Text
            oTable.Cell(row, 3).Range.Text = "[g/Am3]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Efficiency"
            oTable.Cell(row, 2).Range.Text = TextBox21.Text
            oTable.Cell(row, 3).Range.Text = "[&]"

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
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 24, 10)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Calc.data"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Dia class[mu]"
            oTable.Cell(row, 2).Range.Text = "Feed psd cumm [%]"
            oTable.Cell(row, 3).Range.Text = "Feed psd diff [%]"
            oTable.Cell(row, 4).Range.Text = "Loss of feed [%]"
            oTable.Cell(row, 5).Range.Text = "Loss abs [%]"
            oTable.Cell(row, 6).Range.Text = "Loss cum [%]"
            oTable.Cell(row, 7).Range.Text = "Catch abs [%]"
            oTable.Cell(row, 8).Range.Text = "Catch cum [%]"
            oTable.Cell(row, 9).Range.Text = "Efficiency [%]"

            For j = 0 To 22
                row += 1
                oTable.Cell(row, 1).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(0).Value, String)
                oTable.Cell(row, 2).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(1).Value, String)
                oTable.Cell(row, 3).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(2).Value, String)
                oTable.Cell(row, 4).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(3).Value, String)
                oTable.Cell(row, 5).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(4).Value, String)
                oTable.Cell(row, 6).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(5).Value, String)
                oTable.Cell(row, 7).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(6).Value, String)
                oTable.Cell(row, 8).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(7).Value, String)
                oTable.Cell(row, 9).Range.Text = CType(DataGridView1.Rows.Item(j).Cells(8).Value, String)
            Next

            For j = 1 To 8
                oTable.Columns(j).Width = oWord.InchesToPoints(0.8)   'Change width of columns 
            Next
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------save Chart2 (Loss curve)---------------- 
            Draw_chart2(Chart2)
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

    Private Sub Present_loss_grid1()
        Dim ks As Integer   'present Case
        Dim j As Integer
        Dim total_abs_loss_C As Double = 0
        Dim total_abs_loss As Double = 0
        Dim total_psd_diff As Double = 0

        ks = CInt(NumericUpDown30.Value)   'present Case
        DataGridView2.ColumnCount = 18
        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add(111)
        'DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        DataGridView2.Columns(0).HeaderText = "Dia class [mu]"
        DataGridView2.Columns(1).HeaderText = "Dia average [mu]"
        DataGridView2.Columns(2).HeaderText = "Dia/k [-]"
        DataGridView2.Columns(3).HeaderText = "Loss overall [-]"
        DataGridView2.Columns(4).HeaderText = "Loss overall Corrected"
        DataGridView2.Columns(5).HeaderText = "Catch chart [%]"     '
        DataGridView2.Columns(6).HeaderText = "Group number [-]"    '
        DataGridView2.Columns(7).HeaderText = "d1 lower dia [mu]"   '
        DataGridView2.Columns(8).HeaderText = "d2 upper dia [mu]"   '
        DataGridView2.Columns(9).HeaderText = "p1 input [%]"        '
        DataGridView2.Columns(10).HeaderText = "p2 input [%]"       '
        DataGridView2.Columns(11).HeaderText = "k [-]"              '    
        DataGridView2.Columns(12).HeaderText = "m [-]"              '

        DataGridView2.Columns(13).HeaderText = "i_psd cum [-]"      '
        DataGridView2.Columns(14).HeaderText = "psd cum [%]"        '
        DataGridView2.Columns(15).HeaderText = "psd diff [%]"       '
        DataGridView2.Columns(16).HeaderText = "loss abs [%]"       '
        DataGridView2.Columns(17).HeaderText = "loss corr abs [%]"  '

        For row = 1 To 110  'Fill the DataGrid
            j = row - 1
            DataGridView2.Rows.Item(j).Cells(0).Value = _cees(ks).stage1(j).dia.ToString("F5")
            DataGridView2.Rows.Item(j).Cells(1).Value = _cees(ks).stage1(j).d_ave.ToString("F5")     'Average diameter
            DataGridView2.Rows.Item(j).Cells(2).Value = _cees(ks).stage1(j).d_ave_K.ToString("F5")   'Average dia/K stokes
            DataGridView2.Rows.Item(j).Cells(3).Value = _cees(ks).stage1(j).loss_overall.ToString("F5")   'Loss 
            DataGridView2.Rows.Item(j).Cells(4).Value = _cees(ks).stage1(j).loss_overall_C.ToString("F5") 'Loss 
            DataGridView2.Rows.Item(j).Cells(5).Value = _cees(ks).stage1(j).catch_chart.ToString("F5")    'Catch
            DataGridView2.Rows.Item(j).Cells(6).Value = _cees(ks).stage1(j).i_grp.ToString      'Groep nummer
            DataGridView2.Rows.Item(j).Cells(7).Value = _cees(ks).stage1(j).i_d1.ToString("F5")       'class lower dia limit
            DataGridView2.Rows.Item(j).Cells(8).Value = _cees(ks).stage1(j).i_d2.ToString("F5")       'class upper dia limit
            DataGridView2.Rows.Item(j).Cells(9).Value = _cees(ks).stage1(j).i_p1.ToString("F5")     'User input percentage
            DataGridView2.Rows.Item(j).Cells(10).Value = _cees(ks).stage1(j).i_p2.ToString("F5")    'User input percentage
            DataGridView2.Rows.Item(j).Cells(11).Value = _cees(ks).stage1(j).i_k.ToString("F3")   '
            DataGridView2.Rows.Item(j).Cells(12).Value = _cees(ks).stage1(j).i_m.ToString("F5")    '
            DataGridView2.Rows.Item(j).Cells(13).Value = _cees(ks).stage1(j).psd_cum.ToString("F5")  '
            DataGridView2.Rows.Item(j).Cells(14).Value = _cees(ks).stage1(j).psd_cum_pro.ToString("F3")   '[%]
            DataGridView2.Rows.Item(j).Cells(15).Value = _cees(ks).stage1(j).psd_dif.ToString("F3")     '[%]
            DataGridView2.Rows.Item(j).Cells(16).Value = _cees(ks).stage1(j).loss_abs.ToString("F3")    '[%]
            DataGridView2.Rows.Item(j).Cells(17).Value = _cees(ks).stage1(j).loss_abs_C.ToString("F3")  '[%]
            total_psd_diff += _cees(ks).stage1(j).psd_dif
            total_abs_loss += _cees(ks).stage1(j).loss_abs
            total_abs_loss_C += _cees(ks).stage1(j).loss_abs_C
        Next
        DataGridView2.Rows.Item(111).Cells(15).Value = total_psd_diff.ToString("F5")
        DataGridView2.Rows.Item(111).Cells(16).Value = total_abs_loss.ToString("F5")
        DataGridView2.Rows.Item(111).Cells(17).Value = total_abs_loss_C.ToString("F5")
    End Sub

    Private Sub Present_loss_grid2()
        Dim j As Integer
        Dim ks As Integer 'Present case number
        Dim total_abs_loss_C As Double = 0
        Dim total_abs_loss As Double = 0
        Dim total_psd_diff1 As Double = 0
        Dim total_psd_diff2 As Double = 0

        ks = CInt(NumericUpDown30.Value)

        DataGridView3.ColumnCount = 19
        DataGridView3.Rows.Clear()
        DataGridView3.Rows.Add(111)
        'DatagridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        DataGridView3.Columns(0).HeaderText = "Dia class [mu]"
        DataGridView3.Columns(1).HeaderText = "Dia average [mu]"
        DataGridView3.Columns(2).HeaderText = "Dia/k [-]"
        DataGridView3.Columns(3).HeaderText = "Loss overall [-]"
        DataGridView3.Columns(4).HeaderText = "Loss overall Corrected [-]"
        DataGridView3.Columns(5).HeaderText = "Catch chart [%]"     '
        DataGridView3.Columns(6).HeaderText = "Group number"    '
        DataGridView3.Columns(7).HeaderText = "d1 lower dia [mu]"    '
        DataGridView3.Columns(8).HeaderText = "d2 upper dia [mu]"    '
        DataGridView3.Columns(9).HeaderText = "p1 input [%]"        '
        DataGridView3.Columns(10).HeaderText = "p2 input [%]"       '
        DataGridView3.Columns(11).HeaderText = "k [-]"              '    
        DataGridView3.Columns(12).HeaderText = "m [-]"              '

        DataGridView3.Columns(13).HeaderText = "i_psd cum [-]"      '
        DataGridView3.Columns(14).HeaderText = "psd cum [%] for chart"    '
        DataGridView3.Columns(15).HeaderText = "psd diff [%] of 1-stage"       '
        DataGridView3.Columns(16).HeaderText = "psd diff [%] of 2-stage"       '
        DataGridView3.Columns(17).HeaderText = "loss abs [%]"   '
        DataGridView3.Columns(18).HeaderText = "loss corr abs [%]" '

        For row = 1 To 110  'Fill the DataGrid
            j = row - 1
            DataGridView3.Rows.Item(j).Cells(0).Value = _cees(ks).stage2(j).dia.ToString("F5")          'Dia particle
            DataGridView3.Rows.Item(j).Cells(1).Value = _cees(ks).stage2(j).d_ave.ToString("F5")        'Average diameter
            DataGridView3.Rows.Item(j).Cells(2).Value = _cees(ks).stage2(j).d_ave_K.ToString("F5")      'Average dia/K stokes
            DataGridView3.Rows.Item(j).Cells(3).Value = _cees(ks).stage2(j).loss_overall.ToString("F5")  'Loss 
            DataGridView3.Rows.Item(j).Cells(4).Value = _cees(ks).stage2(j).loss_overall_C.ToString("F5") 'Loss 
            DataGridView3.Rows.Item(j).Cells(5).Value = _cees(ks).stage2(j).catch_chart.ToString("F5")  'Catch
            DataGridView3.Rows.Item(j).Cells(6).Value = _cees(ks).stage2(j).i_grp.ToString              'Groep nummer
            DataGridView3.Rows.Item(j).Cells(7).Value = _cees(ks).stage2(j).i_d1.ToString("F5")         'class lower dia limit
            DataGridView3.Rows.Item(j).Cells(8).Value = _cees(ks).stage2(j).i_d2.ToString("F5")         'class upper dia limit
            DataGridView3.Rows.Item(j).Cells(9).Value = _cees(ks).stage2(j).i_p1.ToString("F5")         'User input percentage
            DataGridView3.Rows.Item(j).Cells(10).Value = _cees(ks).stage2(j).i_p2.ToString("F5")        '
            DataGridView3.Rows.Item(j).Cells(11).Value = _cees(ks).stage2(j).i_k.ToString("F5")         'User input percentage
            DataGridView3.Rows.Item(j).Cells(12).Value = _cees(ks).stage2(j).i_m.ToString("F5")         '
            DataGridView3.Rows.Item(j).Cells(13).Value = _cees(ks).stage2(j).psd_cum.ToString("F5")     '
            DataGridView3.Rows.Item(j).Cells(14).Value = _cees(ks).stage2(j).psd_cum_pro.ToString("F5")    '[%]
            DataGridView3.Rows.Item(j).Cells(15).Value = _cees(ks).stage1(j).psd_dif.ToString("F5")     '[%] (stage 1 !!!!)
            DataGridView3.Rows.Item(j).Cells(16).Value = _cees(ks).stage2(j).psd_dif.ToString("F5")     '[%]
            DataGridView3.Rows.Item(j).Cells(17).Value = _cees(ks).stage2(j).loss_abs.ToString("F5")    '[%]
            DataGridView3.Rows.Item(j).Cells(18).Value = _cees(ks).stage2(j).loss_abs_C.ToString("F5")  '[%]
            total_psd_diff1 += _cees(ks).stage1(j).psd_dif
            total_psd_diff2 += _cees(ks).stage2(j).psd_dif
            total_abs_loss += _cees(ks).stage2(j).loss_abs
            total_abs_loss_C += _cees(ks).stage2(j).loss_abs_C
        Next
        DataGridView3.Rows.Item(111).Cells(15).Value = total_psd_diff1.ToString("F5")
        DataGridView3.Rows.Item(111).Cells(16).Value = total_psd_diff2.ToString("F5")
        DataGridView3.Rows.Item(111).Cells(17).Value = total_abs_loss.ToString("F5")
        DataGridView3.Rows.Item(111).Cells(18).Value = total_abs_loss_C.ToString("F5")
    End Sub

    Private Sub Calc_k_and_m(ByRef g As GvG_Calc_struct)
        g.i_k = Log(Log(g.i_p1) / Log(g.i_p2)) / Log(g.i_d1 / g.i_d2)   '====== k ===========
        g.i_m = g.i_d1 / ((-Log(g.i_p1)) ^ (1 / g.i_k))                 '====== m ===========
    End Sub

    Private Sub Calc_stage1(ks As Integer)
        'This is the standard VTK cyclone calculation for case "ks" 

        Dim i As Integer = 0
        Dim dia_max As Double       'Above this diameter everything is caught
        Dim dia_min As Double       'Below this diameter nothing is caught
        Dim istep As Double         'Particle diameter step
        Dim perc_smallest_part1 As Double
        Dim fac_m As Double
        Dim words() As String

        '------ the idea is that the smallest diameter cyclone determines
        '------ the smallest particle diameter used in the calculation
        '------ for the stage #1 cyclone
        If numericUpDown5.Value > NumericUpDown35.Value Then
            _cees(ks).stage1(0).dia = Calc_dia_particle(1.0, _cees(ks).Kstokes2, 2) 'stage #2 cyclone
        Else
            _cees(ks).stage1(0).dia = Calc_dia_particle(1.0, _cees(ks).Kstokes1, 1) 'stage #1 cyclone
        End If

        _cees(ks).stage1(0).d_ave = _cees(ks).stage1(0).dia / 2                       'Average diameter
        _cees(ks).stage1(0).d_ave_K = _cees(ks).stage1(0).d_ave / _cees(ks).Kstokes1  'dia/k_stokes
        _cees(ks).stage1(0).loss_overall = Calc_verlies(_cees(ks).stage1(0).d_ave_K, False, _cees(ks).Kstokes1, 1)     '[-] loss overall
        Calc_verlies_corrected(_cees(ks).stage1(0), 1)                                '[-] loss overall corrected
        _cees(ks).stage1(0).catch_chart = (1 - _cees(ks).stage1(i).loss_overall_C) * 100     '[%]

        Size_classification(_cees(ks).stage1(0))                                    'Classify this part size
        Calc_k_and_m(_cees(ks).stage1(0))

        'TextBox24.Text &= "stage1(0).dia=" & _cees(ks).stage1(0).dia.ToString
        'TextBox24.Text &= "   stage1(0).i_d1=" & _cees(ks).stage1(0).i_d1.ToString
        'TextBox24.Text &= ",  stage1(0).i_d2=" & _cees(ks).stage1(0).i_d2.ToString
        'TextBox24.Text &= ",  stage1(0).i_p1=" & _cees(ks).stage1(0).i_p1.ToString
        'TextBox24.Text &= ",  stage1(0).i_p2=" & _cees(ks).stage1(0).i_p1.ToString
        'TextBox24.Text &= ",  stage1(0).i_k=" & _cees(ks).stage1(0).i_k.ToString
        'TextBox24.Text &= ",  stage1(0).i_grp=" & _cees(ks).stage1(0).i_grp.ToString & vbCrLf

        _cees(ks).stage1(0).psd_cum = Math.E ^ (-((_cees(ks).stage1(i).dia / _cees(ks).stage1(i).i_m) ^ _cees(ks).stage1(i).i_k))
        _cees(ks).stage1(0).psd_cum_pro = _cees(ks).stage1(i).psd_cum * 100

        _cees(ks).stage1(0).psd_dif = 100 * (1 - _cees(ks).stage1(i).psd_cum)
        _cees(ks).stage1(0).loss_abs = _cees(ks).stage1(i).loss_overall * _cees(ks).stage1(i).psd_dif
        _cees(ks).stage1(0).loss_abs_C = _cees(ks).stage1(i).loss_overall_C * _cees(ks).stage1(i).psd_dif

        _cees(ks).sum_psd_diff1 = _cees(ks).stage1(0).psd_dif
        _cees(ks).sum_loss1 = _cees(ks).stage1(0).loss_abs
        _cees(ks).sum_loss_C1 = _cees(ks).stage1(0).loss_abs_C

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

        perc_smallest_part1 = 0.0000001                      'smallest particle [%]
        _cees(ks).Dmax1 = Calc_dia_particle(perc_smallest_part1, _cees(ks).Kstokes1, 1)     '=100% loss (biggest particle)
        _cees(ks).Dmin1 = _cees(ks).Kstokes1 * fac_m        'diameter smallest particle caught

        dia_min = CDbl(IIf(_cees(ks).Dmin1 < _cees(ks).Dmin2, _cees(ks).Dmin1, _cees(ks).Dmin2))     'smalles particle
        dia_max = CDbl(IIf(_cees(ks).Dmax1 > _cees(ks).Dmax2, _cees(ks).Dmax1, _cees(ks).Dmax2))     'biggest particle

        '------------ Particle diameter calculation step -----
        istep = (dia_max / dia_min) ^ (1 / 110)             'Calculation step

        For i = 1 To 110
            _cees(ks).stage1(i).dia = _cees(ks).stage1(i - 1).dia * istep
            _cees(ks).stage1(i).d_ave = (_cees(ks).stage1(i - 1).dia + _cees(ks).stage1(i).dia) / 2       'Average diameter
            _cees(ks).stage1(i).d_ave_K = _cees(ks).stage1(i).d_ave / _cees(ks).Kstokes1        'dia/k_stokes
            _cees(ks).stage1(i).loss_overall = Calc_verlies(_cees(ks).stage1(i).d_ave, False, _cees(ks).Kstokes1, 1)   '[-] loss overall
            Calc_verlies_corrected(_cees(ks).stage1(i), 1)                               '[-] loss overall corrected

            If CheckBox2.Checked Then
                _cees(ks).stage1(i).catch_chart = (1 - _cees(ks).stage1(i).loss_overall_C) * 100 '[%] Corrected
            Else
                _cees(ks).stage1(i).catch_chart = (1 - _cees(ks).stage1(i).loss_overall) * 100    '[%] NOT corrected
            End If
            Size_classification(_cees(ks).stage1(i))                                   'Classify this part size

            If _cees(ks).stage1(i).i_grp <> 11 And _cees(ks).stage2(i).i_grp <> 11 Then 'to prevent silly results
                Calc_k_and_m(_cees(ks).stage1(i))
                _cees(ks).stage1(i).psd_cum = Math.E ^ (-((_cees(ks).stage1(i).dia / _cees(ks).stage1(i).i_m) ^ _cees(ks).stage1(i).i_k))
                _cees(ks).stage1(i).psd_cum_pro = _cees(ks).stage1(i).psd_cum * 100
                _cees(ks).stage1(i).psd_dif = 100 * (_cees(ks).stage1(i - 1).psd_cum - _cees(ks).stage1(i).psd_cum)
            Else
                _cees(ks).stage1(i).i_k = 0
                _cees(ks).stage1(i).i_m = 0
                _cees(ks).stage1(i).psd_cum = 0
                _cees(ks).stage1(i).psd_cum_pro = 0
                _cees(ks).stage1(i).psd_dif = 0
            End If
            _cees(ks).stage1(i).loss_abs = _cees(ks).stage1(i).loss_overall * _cees(ks).stage1(i).psd_dif
            _cees(ks).stage1(i).loss_abs_C = _cees(ks).stage1(i).loss_overall_C * _cees(ks).stage1(i).psd_dif

            '----- sum value -----
            _cees(ks).sum_psd_diff1 += _cees(ks).stage1(i).psd_dif
            _cees(ks).sum_loss1 += _cees(ks).stage1(i).loss_abs
            _cees(ks).sum_loss_C1 += _cees(ks).stage1(i).loss_abs_C
        Next
        _cees(ks).loss_total1 = _cees(ks).sum_loss_C1 + ((100 - _cees(ks).sum_psd_diff1) * perc_smallest_part1)

        If CheckBox4.Checked Then   'Stage #1 is bypassed
            _cees(ks).emmis1 = NumericUpDown4.Value '[g/Am3]
        Else                        'Stage #1 is not bypassed
            _cees(ks).emmis1 = NumericUpDown4.Value * _cees(ks).loss_total1 / 100  '[g/Am3]
        End If

        _cees(ks).stofb2 = _cees(ks).emmis1 'Dust load stage #2 in emission stage #1
        CheckBox3.Checked = CBool(IIf(_cees(ks).stofb2 > 5, True, False))
        _cees(ks).Efficiency1 = 100 - _cees(ks).loss_total1        '[%] Efficiency

        '----------- present -----------
        TextBox51.Text = dia_max.ToString("F1")             'diameter [mu] 100% catch
        TextBox52.Text = dia_min.ToString("F2")             'diameter [mu] 100% loss
        TextBox56.Text = ComboBox1.Text                     'Cyclone typr
        TextBox57.Text = CheckBox2.Checked.ToString         'Correction 
        TextBox70.Text = _cees(ks).stofb2.ToString("F2")    'Dust load


        TextBox118.Text = _cees(ks).sum_psd_diff1.ToString("F3")
        TextBox54.Text = _cees(ks).sum_loss1.ToString("F3")
        TextBox34.Text = _cees(ks).sum_loss_C1.ToString("F3")

        'If CheckBox2.Checked Then
        TextBox58.Text = _cees(ks).loss_total1.ToString("F5")    'Corrected ??????
            TextBox59.Text = _cees(ks).Efficiency1.ToString("F3")
            TextBox21.Text = TextBox59.Text
            TextBox60.Text = _cees(ks).emmis1.ToString("F3")
            TextBox18.Text = TextBox60.Text
        'Else
        '    TextBox58.Text = _cees(ks).sum_loss1.ToString("F5")      'NOT Corrected  ??????
        '    TextBox59.Text = _cees(ks).Efficiency1.ToString("F3")
        '    TextBox21.Text = TextBox59.Text
        '    TextBox60.Text = (NumericUpDown4.Value * _cees(ks).sum_loss1 / 100).ToString("0.000")
        '    TextBox18.Text = TextBox60.Text
        'End If
    End Sub

    Private Sub Calc_stage2(ks As Integer)
        'This is the standard VTK cyclone calculation 
        Dim i As Integer = 0
        Dim dia_max As Double       'Above this diameter everything is caught
        Dim dia_min As Double       'Below this diameter nothing is caught
        Dim istep As Double         'Particle diameter step
        Dim perc_smallest_part2 As Double
        Dim fac_m As Double
        Dim words() As String
        Dim kgh, tot_kgh As Double

        '----------- stof belasting ------------
        tot_kgh = _cees(ks).Flow2 * _cees(ks).stofb2 / 1000 * 3600 * _cees(ks).Noc2     '[kg/hr] Dust inlet 

        'TextBox24.Text &= "_cees(ks).stofb2= " & _cees(ks).stofb2.ToString
        'TextBox24.Text &= ",  _cees(ks).Flow2= " & (_cees(ks).Flow2 * 3600).ToString & vbCrLf


        kgh = tot_kgh / _cees(ks).Noc2                          '[kg/hr/Cy] Dust inlet 
        TextBox100.Text = kgh.ToString("0")
        TextBox101.Text = tot_kgh.ToString("0")

        '--------- now the particles (====Grid line 0======)------------
        _cees(ks).stage2(i).dia = _cees(ks).stage1(i).dia                                   'Copy stage #1
        _cees(ks).stage2(i).d_ave = _cees(ks).stage2(0).dia / 2                             'Average diameter
        _cees(ks).stage2(i).d_ave_K = _cees(ks).stage2(0).d_ave / _cees(ks).Kstokes2        'dia/k_stokes
        _cees(ks).stage2(i).loss_overall = Calc_verlies(_cees(ks).stage2(0).d_ave_K, False, _cees(ks).Kstokes2, 2)     '[-] loss overall
        Calc_verlies_corrected(_cees(ks).stage2(0), 2)                               '[-] loss overall corrected
        _cees(ks).stage2(i).catch_chart = (1 - _cees(ks).stage2(i).loss_overall_C) * 100    '[%]
        Size_classification(_cees(ks).stage2(i))                                  'groepnummer

        Calc_k_and_m(_cees(ks).stage2(i))
        _cees(ks).stage2(i).psd_cum = Math.E ^ (-((_cees(ks).stage2(i).dia / _cees(ks).stage2(i).i_m) ^ _cees(ks).stage2(i).i_k))
        _cees(ks).stage2(i).psd_cum_pro = _cees(ks).stage2(i).psd_cum * 100
        _cees(ks).stage2(i).psd_dif = 100 * _cees(ks).stage1(i).loss_abs / (100 - _cees(ks).Efficiency1)              'LOSS STAGE #1
        _cees(ks).stage2(i).loss_abs = _cees(ks).stage1(i).loss_overall * _cees(ks).stage2(i).psd_dif
        _cees(ks).stage2(i).loss_abs_C = _cees(ks).stage1(i).loss_overall_C * _cees(ks).stage2(i).psd_dif

        TextBox24.Text &= "_cees(ks).Efficiency1= " & _cees(ks).Efficiency1.ToString & vbCrLf

        _cees(ks).sum_psd_diff2 = _cees(ks).stage2(i).psd_dif
        _cees(ks).sum_loss2 = _cees(ks).stage2(i).loss_abs
        _cees(ks).sum_loss_C2 = _cees(ks).stage2(i).loss_abs_C

        '------ increment step --------
        'stapgrootte bij 110-staps logaritmische verdeling van het
        'deeltjesdiameter-bereik van loss=100% tot 0,00000001%
        'Deze wordt gebruikt voor het opstellen van de gefractioneerde
        'verliescurve.

        '-------------- korrelgrootte factoren ------
        If ComboBox2.SelectedIndex > -1 Then
            words = rekenlijnen(ComboBox2.SelectedIndex).Split(CType(";", Char()))
            '---- diameter kleiner dan dia kritisch
            fac_m = CDbl(words(2))
        End If

        '------------ Find the biggest and smallest particle -----
        perc_smallest_part2 = 0.0000001                      'smallest particle [%]
        _cees(ks).Dmax2 = Calc_dia_particle(perc_smallest_part2, _cees(ks).Kstokes2, 2)     '=100% loss (biggest particle)
        _cees(ks).Dmin2 = _cees(ks).Kstokes2 * fac_m                'diameter smallest particle caught

        dia_min = CDbl(IIf(_cees(ks).Dmin1 < _cees(ks).Dmin2, _cees(ks).Dmin1, _cees(ks).Dmin2))     '=100% loss (biggest particle)
        dia_max = CDbl(IIf(_cees(ks).Dmax1 > _cees(ks).Dmax2, _cees(ks).Dmax1, _cees(ks).Dmax2))     '=100% loss (biggest particle)

        '------------ Particle diameter calculation step -----
        istep = (dia_max / dia_min) ^ (1 / 110)             'Calculation step

        'TextBox24.Text &= "dia_min= " & dia_min.ToString
        'TextBox24.Text &= ",  dia_max= " & dia_max.ToString
        'TextBox24.Text &= ",  istep= " & istep.ToString & vbCrLf


        For i = 1 To 110    '=========Stage #2, Grid lines 1...============ 
            _cees(ks).stage2(i).dia = _cees(ks).stage1(i).dia        'Diameter Copy stage #1
            _cees(ks).stage2(i).d_ave = _cees(ks).stage1(i).d_ave            'Average diameter
            _cees(ks).stage2(i).d_ave_K = _cees(ks).stage2(i).d_ave / _cees(ks).Kstokes2          'dia/k_stokes
            _cees(ks).stage2(i).loss_overall = Calc_verlies(_cees(ks).stage2(i).d_ave, False, _cees(ks).Kstokes2, 2)   '[-] loss overall
            Calc_verlies_corrected(_cees(ks).stage2(i), 2)                                '[-] loss overall corrected
            If CheckBox3.Checked Then
                _cees(ks).stage2(i).catch_chart = (1 - _cees(ks).stage2(i).loss_overall_C) * 100  '[%] Corrected
            Else
                _cees(ks).stage2(i).catch_chart = (1 - _cees(ks).stage2(i).loss_overall) * 100    '[%] NOT corrected
            End If
            Size_classification(_cees(ks).stage2(i))                                     'Calc

            If _cees(ks).stage1(i).i_grp <> 11 And _cees(ks).stage2(i).i_grp <> 11 Then 'to prevent silly results 
                Calc_k_and_m(_cees(ks).stage2(i))
                _cees(ks).stage2(i).psd_cum = Math.E ^ (-((_cees(ks).stage2(i).dia / _cees(ks).stage2(i).i_m) ^ _cees(ks).stage2(i).i_k))
                _cees(ks).stage2(i).psd_cum_pro = _cees(ks).stage2(i).psd_cum * 100
                _cees(ks).stage2(i).psd_dif = 100 * (_cees(ks).stage2(i - 1).psd_cum - _cees(ks).stage2(i).psd_cum)
            Else
                _cees(ks).stage2(i).i_k = 0
                _cees(ks).stage2(i).i_m = 0
                _cees(ks).stage2(i).psd_cum = 0
                _cees(ks).stage2(i).psd_cum_pro = 0
                _cees(ks).stage2(i).psd_dif = 0
            End If
            _cees(ks).stage2(i).loss_abs = _cees(ks).stage1(i).loss_overall * _cees(ks).stage2(i).psd_dif
            _cees(ks).stage2(i).loss_abs_C = _cees(ks).stage1(i).loss_overall_C * _cees(ks).stage2(i).psd_dif

            '----- sum value -----
            _cees(ks).sum_psd_diff2 += _cees(ks).stage2(i).psd_dif
            _cees(ks).sum_loss2 += _cees(ks).stage2(i).loss_abs
            _cees(ks).sum_loss_C2 += _cees(ks).stage2(i).loss_abs_C
        Next
        _cees(ks).loss_total2 = _cees(ks).sum_loss_C2 + ((100 - _cees(ks).sum_psd_diff2) * perc_smallest_part2)
        _cees(ks).emmis2 = _cees(ks).emmis1 * _cees(ks).loss_total2 / 100
        _cees(ks).Efficiency2 = 100 - _cees(ks).loss_total2      '[%] Efficiency

        '----------- present -----------
        TextBox63.Text = ComboBox2.Text                 'Cyclone type
        TextBox64.Text = CheckBox3.Checked.ToString     'Hi load correction
        TextBox110.Text = dia_max.ToString("F1")        'diameter [mu] 100% catch
        TextBox111.Text = dia_min.ToString("F2")        'diameter [mu] 100% loss
        TextBox116.Text = istep.ToString("F5")          'Calculation step

        TextBox117.Text = _cees(ks).sum_psd_diff2.ToString("F3")
        TextBox68.Text = _cees(ks).sum_loss2.ToString("F3")
        TextBox69.Text = _cees(ks).sum_loss_C2.ToString("F3")


        'If CheckBox3.Checked Then
        TextBox65.Text = _cees(ks).loss_total2.ToString("F5")    'Corrected
            TextBox66.Text = _cees(ks).Efficiency2.ToString("F3")
            TextBox109.Text = _cees(ks).Efficiency2.ToString("F3")
            TextBox62.Text = _cees(ks).emmis2.ToString("F3")
            'Else
            '    TextBox65.Text = _cees(ks).sum_loss2.ToString("F5")      'NOT Corrected
            '    TextBox66.Text = _cees(ks).Efficiency2.ToString("F3")
            '    TextBox109.Text = _cees(ks).Efficiency2.ToString("F3")
            '    TextBox62.Text = _cees(ks).emmis2.ToString("F3")

            'End If
            TextBox108.Text = TextBox62.Text
    End Sub
    'Determine the particle diameter class upper and lower limits
    ' Private Function Size_classification(dia As Double, noi As Integer) As Double
    Public Sub Size_classification(ByRef g As GvG_Calc_struct)
        If g.dia > 0 Then
            Select Case True
                Case g.dia < NumericUpDown15.Value      '0-10 mu
                    g.i_d1 = NumericUpDown15.Value      '10 mu
                    g.i_d2 = NumericUpDown23.Value      '15 mu
                    g.i_p1 = numericUpDown6.Value / 100
                    g.i_p1 = numericUpDown7.Value / 100
                    g.i_grp = 0
                Case g.dia >= NumericUpDown15.Value And g.dia < NumericUpDown23.Value   '>=10 and < 15
                    g.i_d1 = NumericUpDown15.Value      '10 mu
                    g.i_d2 = NumericUpDown23.Value      '15 mu
                    g.i_p1 = numericUpDown6.Value / 100
                    g.i_p2 = numericUpDown7.Value / 100
                    g.i_grp = 1
                Case g.dia >= NumericUpDown23.Value And g.dia < NumericUpDown24.Value
                    g.i_d1 = NumericUpDown23.Value      '15 mu
                    g.i_d2 = NumericUpDown24.Value      '20 mu
                    g.i_p1 = numericUpDown7.Value / 100
                    g.i_p2 = numericUpDown8.Value / 100
                    g.i_grp = 2
                Case g.dia >= NumericUpDown24.Value And g.dia < NumericUpDown25.Value
                    g.i_d1 = NumericUpDown24.Value '20
                    g.i_d2 = NumericUpDown25.Value '30
                    g.i_p1 = numericUpDown8.Value / 100
                    g.i_p2 = numericUpDown9.Value / 100
                    g.i_grp = 3
                Case g.dia >= NumericUpDown25.Value And g.dia < NumericUpDown26.Value
                    g.i_d1 = NumericUpDown25.Value '30
                    g.i_d2 = NumericUpDown26.Value '40
                    g.i_p1 = numericUpDown9.Value / 100
                    g.i_p2 = numericUpDown10.Value / 100
                    g.i_grp = 4
                Case g.dia >= NumericUpDown26.Value And g.dia < NumericUpDown27.Value
                    g.i_d1 = NumericUpDown26.Value '40
                    g.i_d2 = NumericUpDown27.Value '50
                    g.i_p1 = numericUpDown10.Value / 100
                    g.i_p2 = numericUpDown11.Value / 100
                    g.i_grp = 5
                Case g.dia >= NumericUpDown27.Value And g.dia < NumericUpDown28.Value
                    g.i_d1 = NumericUpDown27.Value   '50
                    g.i_d2 = NumericUpDown28.Value '60
                    g.i_p1 = numericUpDown11.Value / 100
                    g.i_p2 = numericUpDown12.Value / 100
                    g.i_grp = 6
                Case g.dia >= NumericUpDown28.Value And g.dia < NumericUpDown29.Value
                    g.i_d1 = NumericUpDown28.Value '60
                    g.i_d2 = NumericUpDown29.Value '80
                    g.i_p1 = numericUpDown12.Value / 100
                    g.i_p2 = numericUpDown13.Value / 100
                    g.i_grp = 7
                Case g.dia >= NumericUpDown29.Value And g.dia < NumericUpDown35.Value
                    g.i_d1 = NumericUpDown29.Value      '
                    g.i_d2 = NumericUpDown35.Value      '
                    g.i_p1 = numericUpDown13.Value / 100
                    g.i_p2 = NumericUpDown38.Value / 100
                    g.i_grp = 8
                Case g.dia >= NumericUpDown35.Value And g.dia < NumericUpDown36.Value
                    g.i_d1 = NumericUpDown35.Value  '
                    g.i_d2 = NumericUpDown36.Value  '
                    g.i_p1 = NumericUpDown38.Value / 100
                    g.i_p2 = NumericUpDown39.Value / 100
                    g.i_grp = 9
                Case g.dia >= NumericUpDown36.Value And g.dia < NumericUpDown37.Value
                    g.i_d1 = NumericUpDown36.Value  '
                    g.i_d2 = NumericUpDown37.Value  '
                    g.i_p1 = NumericUpDown39.Value / 100
                    g.i_p2 = NumericUpDown40.Value / 100
                    g.i_grp = 10
                Case Else
                    g.i_d1 = NumericUpDown37.Value  '
                    g.i_d2 = g.i_d1 * 1.0001
                    g.i_p1 = NumericUpDown40.Value / 100
                    g.i_p2 = 1.0
                    g.i_grp = 11
            End Select

            Dim w(11) As Double  'Individual particle class weights
            w(0) = NumericUpDown40.Value
            w(1) = NumericUpDown39.Value - w(0)
            w(2) = NumericUpDown38.Value - w(1) - w(0)
            w(3) = numericUpDown13.Value - w(2) - w(1) - w(0)
            w(4) = numericUpDown12.Value - w(3) - w(2) - w(1) - w(0)
            w(5) = numericUpDown11.Value - w(4) - w(3) - w(2) - w(1) - w(0)
            w(6) = numericUpDown10.Value - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(7) = numericUpDown9.Value - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(8) = numericUpDown8.Value - w(7) - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(9) = numericUpDown7.Value - w(8) - w(7) - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(10) = numericUpDown6.Value - w(9) - w(8) - w(7) - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)

            TextBox50.Text = w(10).ToString("0.0")
            TextBox49.Text = w(9).ToString("0.0")
            TextBox46.Text = w(8).ToString("0.0")
            TextBox45.Text = w(7).ToString("0.0")
            TextBox44.Text = w(6).ToString("0.0")
            TextBox43.Text = w(5).ToString("0.0")
            TextBox27.Text = w(4).ToString("0.0")
            TextBox25.Text = w(3).ToString("0.0")
            TextBox81.Text = w(2).ToString("0.0")
            TextBox82.Text = w(1).ToString("0.0")

            '-------- Check -- bigger diameter must have bigger cummulative weight
            NumericUpDown15.BackColor = CType(IIf(NumericUpDown15.Value > 0, Color.LightGreen, Color.Red), Color)
            NumericUpDown23.BackColor = CType(IIf(NumericUpDown23.Value >= NumericUpDown15.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown24.BackColor = CType(IIf(NumericUpDown24.Value >= NumericUpDown23.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown25.BackColor = CType(IIf(NumericUpDown25.Value >= NumericUpDown24.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown26.BackColor = CType(IIf(NumericUpDown26.Value >= NumericUpDown25.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown27.BackColor = CType(IIf(NumericUpDown27.Value >= NumericUpDown26.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown28.BackColor = CType(IIf(NumericUpDown28.Value >= NumericUpDown27.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown29.BackColor = CType(IIf(NumericUpDown29.Value >= NumericUpDown28.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown35.BackColor = CType(IIf(NumericUpDown35.Value >= NumericUpDown29.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown36.BackColor = CType(IIf(NumericUpDown36.Value >= NumericUpDown35.Value, Color.LightGreen, Color.Red), Color)
            NumericUpDown37.BackColor = CType(IIf(NumericUpDown37.Value >= NumericUpDown36.Value, Color.LightGreen, Color.Red), Color)
        Else
            MessageBox.Show("Error in line 1946")
        End If
    End Sub

    'Calculate cyclone weight
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabPage7.Enter, NumericUpDown32.ValueChanged, NumericUpDown31.ValueChanged
        Dim w1, w2, w3, w4, w5, w6 As Double
        Dim c_weight1 As Double             '[kg] cyclone weight
        Dim C_weight2 As Double             '[kg] cyclone weight
        Dim total_instal As Double          '[kg] installation weight
        Dim plt_body1, plt_top1 As Double
        Dim plt_body2, plt_top2 As Double
        Dim ro_steel As Double = 7850       'Density steel
        Dim hh, hj, hk As Double            'Dimensions
        Dim _db As Double
        Dim sheet_metal_wht As Double
        Dim p304, p316 As Double

        '========== Stage cyclone #1 =====
        plt_top1 = NumericUpDown32.Value        '[mm] top plate
        plt_body1 = NumericUpDown31.Value       '[mm] rest of the cyclone
        _db = numericUpDown5.Value / 1000       '[m] Body diameter

        'weight top plate
        w1 = PI / 4 * _db ^ 2 * plt_top1 / 1000 * ro_steel

        'weight cylindrical body
        hh = _cyl1_dim(10) * _db / 1000                  '[m] Length romp
        w2 = PI * _db * hh * plt_body1 * ro_steel        '[kg] weight romp

        'weight cone
        hh = _cyl1_dim(11) * _db / 1000                     '[m] Length cone
        hj = _db                                            '[m] grote diameter cone 
        hk = _cyl1_dim(12) * _db / 1000                     '[m] kleine diameter cone 
        w3 = PI * (hj + hk) / 2 * hh * plt_body1 * ro_steel '[kg] weight cone

        'weight gas outlet pipe
        hh = _cyl1_dim(8) * _db / 1000                      '[m] Length insteekpijp
        hj = _cyl1_dim(7) * _db / 1000                      '[m] Uitlaat flensdiameter inw.
        w4 = PI * hh * hj * plt_body1 * ro_steel            '[kg] weight insteekpijp

        'weight 3P pipe
        hh = _cyl1_dim(13) * _db / 1000                     '[m] Length 3P pijp
        hj = _cyl1_dim(12) * _db / 1000                     '[m] diameter 3P inw.
        w5 = PI * hj * hh * plt_body1 * ro_steel            '[kg] weight 3P pipe

        'weight 3P cone
        hh = _cyl1_dim(14) * _db / 1000                     '[m] Length 3P cone
        hj = _cyl1_dim(12) * _db / 1000                     '[m] grote diameter 3P pijp
        hk = _cyl1_dim(15) * _db / 1000                     '[m] kleine diameter 3P pijp
        w6 = PI * (hj + hk) / 2 * hh * plt_body1 * ro_steel '[kg] weight 3P pipe

        c_weight1 = w1 + w2 + w3 + w4 + w5 + w6             'Total weight
        c_weight1 *= 1.1                                    '10% safety
        TextBox61.Text = c_weight1.ToString("0")            'Total weight


        '========== Stage cyclone #2 ========
        plt_top2 = NumericUpDown41.Value                '[mm] top plate
        plt_body2 = NumericUpDown42.Value               '[mm] rest of the cyclone
        _db = NumericUpDown34.Value / 1000              '[m] Body diameter

        'weight top plate
        w1 = PI / 4 * _db ^ 2 * plt_top2 / 1000 * ro_steel

        'weight cylindrical body
        hh = _cyl1_dim(10) * _db / 1000                     '[m] Length romp
        w2 = PI * _db * hh * plt_body2 * ro_steel           '[kg] weight romp

        'weight cone
        hh = _cyl1_dim(11) * _db / 1000                     '[m] Length cone
        hj = _db                                            '[m] grote diameter cone 
        hk = _cyl1_dim(12) * _db / 1000                     '[m] kleine diameter cone 
        w3 = PI * (hj + hk) / 2 * hh * plt_body2 * ro_steel '[kg] weight cone

        'weight gas outlet pipe
        hh = _cyl1_dim(8) * _db / 1000                      '[m] Length insteekpijp
        hj = _cyl1_dim(7) * _db / 1000                      '[m] Uitlaat flensdiameter inw.
        w4 = PI * hh * hj * plt_body2 * ro_steel            '[kg] weight insteekpijp

        'weight 3P pipe
        hh = _cyl1_dim(13) * _db / 1000                     '[m] Length 3P pijp
        hj = _cyl1_dim(12) * _db / 1000                     '[m] diameter 3P inw.
        w5 = PI * hj * hh * plt_body2 * ro_steel            '[kg] weight 3P pipe

        'weight 3P cone
        hh = _cyl1_dim(14) * _db / 1000                     '[m] Length 3P cone
        hj = _cyl1_dim(12) * _db / 1000                     '[m] grote diameter 3P pijp
        hk = _cyl1_dim(15) * _db / 1000                     '[m] kleine diameter 3P pijp
        w6 = PI * (hj + hk) / 2 * hh * plt_body2 * ro_steel '[kg] weight 3P pipe

        C_weight2 = w1 + w2 + w3 + w4 + w5 + w6             'Total weight
        C_weight2 *= 1.03                                   '3% weight flanges 
        C_weight2 *= 1.1                                    '10% safety

        total_instal = c_weight1 * NumericUpDown20.Value
        total_instal += C_weight2 * NumericUpDown33.Value

        sheet_metal_wht = total_instal * 1.45               'Gross Sheet metal weight
        p304 = sheet_metal_wht * NumericUpDown45.Value      'rvs 304
        p316 = sheet_metal_wht * NumericUpDown44.Value      'rvs 316

        '-------- present -----------
        TextBox72.Text = C_weight2.ToString("0")
        TextBox99.Text = total_instal.ToString("0")
        TextBox83.Text = sheet_metal_wht.ToString("0")
        TextBox114.Text = p304.ToString("0")
        TextBox115.Text = p316.ToString("0")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim c As Integer
        'Save data of screen into the _cees array

        c = CInt(NumericUpDown30.Value)       'Case number
        Fill_cees_array(c)
    End Sub

    Private Sub NumericUpDown30_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown30.ValueChanged
        Case_number_changed()
    End Sub
    Private Sub Case_number_changed()
        Dim zz As Integer = CInt(NumericUpDown30.Value)    'Case number
        Try
            '----------- general (not calculated) data------------------
            TextBox53.Text = _cees(zz).case_name            'Case name
            If _cees(zz).case_name.Length > 0 Then

                '[mu] Class upper particle diameter limit diameter
                NumericUpDown15.Value = CDec(_cees(zz).dia_big(0))   '10
                NumericUpDown23.Value = CDec(_cees(zz).dia_big(1))   '15
                NumericUpDown24.Value = CDec(_cees(zz).dia_big(2))   '20
                NumericUpDown25.Value = CDec(_cees(zz).dia_big(3))   '30
                NumericUpDown26.Value = CDec(_cees(zz).dia_big(4))   '40
                NumericUpDown27.Value = CDec(_cees(zz).dia_big(5))   '50
                NumericUpDown28.Value = CDec(_cees(zz).dia_big(6))   '60
                NumericUpDown29.Value = CDec(_cees(zz).dia_big(7))   '80

                'Percentale van de inlaat stof belasting
                numericUpDown6.Value = CDec(_cees(zz).class_load(0) * 100)
                numericUpDown7.Value = CDec(_cees(zz).class_load(1) * 100)
                numericUpDown8.Value = CDec(_cees(zz).class_load(2) * 100)
                numericUpDown9.Value = CDec(_cees(zz).class_load(3) * 100)
                numericUpDown10.Value = CDec(_cees(zz).class_load(4) * 100)
                numericUpDown11.Value = CDec(_cees(zz).class_load(5) * 100)
                numericUpDown12.Value = CDec(_cees(zz).class_load(6) * 100)
                numericUpDown13.Value = CDec(_cees(zz).class_load(7) * 100)

                NumericUpDown1.Value = CDec(_cees(zz).FlowT)        'Air flow total
                NumericUpDown4.Value = CDec(_cees(zz).stofb1)       'Dust inlet [g/Am3] 
                ComboBox1.SelectedIndex = _cees(zz).Ct1             'Cyclone type stage #1
                ComboBox2.SelectedIndex = _cees(zz).Ct2             'Cyclone type stage #2
                NumericUpDown20.Value = _cees(zz).Noc1              'Cyclone in parallel
                numericUpDown13.Value = CDec(_cees(zz).db1)         'Diameter cyclone body #1
                NumericUpDown34.Value = CDec(_cees(zz).db2)         'Diameter cyclone body #2
                numericUpDown3.Value = CDec(_cees(zz).ro_gas)       'Density [kg/hr]
                numericUpDown2.Value = CDec(_cees(zz).ro_solid)     'Density [kg/hr]
                numericUpDown14.Value = CDec(_cees(zz).visco)       'Visco in Centi Poise
                NumericUpDown18.Value = CDec(_cees(zz).Temp)        'Temperature [c]
                NumericUpDown19.Value = CDec(_cees(zz).Druk1)       'Pressure [mbar]
            End If

        Catch ex As Exception
            'MessageBox.Show(ex.Message & vbcrlf &  "Line 1586")  
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click, TabControl1.Enter
        Calc_sequence()
    End Sub

End Class
