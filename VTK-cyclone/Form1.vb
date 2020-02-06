Imports System.Globalization
Imports System.IO
Imports System.Management
Imports System.Math
Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports Word = Microsoft.Office.Interop.Word
'------- Input data------
'This structure is required for the different operating cases of a cyclone
'Therefore the struct does only contain  the input information
'If the calculation is modified the new result will be found 
<Serializable()> Public Structure Input_struct
    Public Quote_no As String       'Quote number
    Public Tag_no As String         'Tag number
    Public case_name As String      'The case name
    Public Spare_str1 As String     'For future use
    Public Spare_str2 As String     'For future use
    Public Spare_str3 As String     'For future use

    Public FlowT As Double          'Air flow [Am3/h]
    Public dia_big() As Double      'Particle diameter inlet [mu]
    Public class_load() As Double   'group_weight_cum in de inlaat stroom [% weight]
    Public ro_gas As Double         '[kg/hr] Density 
    Public ro_solid As Double       '[kg/hr] Density 
    Public visco As Double          '[Centi Poise] Visco in 
    Public Temp As Double           '[c] Temperature 
    Public Spare1 As String         'For future use
    Public Spare2 As String         'For future use
    Public Spare3 As String         'For future use

    '===== stage #1 parameter ======
    Public Flow1 As Double          '[Am3/s] Air flow per cyclone 
    Public dust1_Am3 As Double      '[g/Am3] Dust load inlet 
    Public dust1_Nm3 As Double      '[g/Nm3] Dust load inlet 
    Public emmis1_Am3 As Double     '[g/Am3] Dust emission 
    Public emmis1_Nm3 As Double     '[g/Nm3] Dust emission 
    Public Efficiency1 As Double    '[%] Efficiency Stage #1 
    Public sum_loss1 As Double      '[-]Passed trough cyclone 
    Public sum_loss_C1 As Double    '[-] Passed trough cyclone Corrected
    Public loss_total1 As Double
    Public sum_psd_diff1 As Double
    Public p1_abs As Double         '[Pa abs] pressure inlet abs
    Public dpgas1 As Double         '[Pa] pressure loss gas
    Public dpdust1 As Double        '[Pa] pressure loss dust
    Public Ro_gas1_Am3 As Double    '[kg/Am3] density gas
    Public Ro_gas1_Nm3 As Double    '[kg/Nm3] density gas
    Public Ct1 As Integer           '[-] Cyclone type (eg AC435)
    Public Noc1 As Integer          '[-] Number paralle Cyclones
    Public db1 As Double            '[m] Diameter cyclone body
    Public inh1 As Double           '[m] inlet hoogte
    Public inb1 As Double           '[m] inlet breedte
    Public dout1 As Double          '[m] diameter zakbuis
    Public inv1 As Double           '[m/s] Inlet velocity cyclone
    Public outv1 As Double          '[m/s] Outlet velocity cyclone
    Public Kstokes1 As Double       'Stokes getal
    Public m1 As Double             'm factor loss curve d< dia critical
    Public stage1() As GvG_Calc_struct   'tbv calculatie stage #1
    Public Dmin1 As Double          'Smallest particle in calculation
    Public Dmax1 As Double          'Biggest particle in calculation

    '===== stage #2 parameters ======
    Public Flow2 As Double          '[Am3/s] Air flow per cyclone 
    Public dust2_Am3 As Double      '[g/Am3] Dust load inlet 
    Public dust2_Nm3 As Double      '[g/Nm3] Dust load inlet 
    Public emmis2_Am3 As Double     '[g/Am3] Dust emission 
    Public emmis2_Nm3 As Double     '[g/Nm3] Dust emission 
    Public sum_loss2 As Double      'Passed trough cyclone
    Public sum_loss_C2 As Double    'Passed trough cyclone Corrected
    Public loss_total2 As Double
    Public sum_psd_diff2 As Double
    Public Efficiency2 As Double    'Efficiency Stage #1 [%}
    Public p2_abs As Double         '[Pa abs] pressure inlet abs
    Public dpgas2 As Double         '[Pa] pressure loss gas
    Public dpdust2 As Double        '[Pa] pressure loss dust
    Public Ro_gas2_Am3 As Double    '[kg/Am3] density gas
    Public Ro_gas2_Nm3 As Double    '[kg/Nm3] density gas
    Public Ct2 As Integer           '[-] Cyclone type (eg AC435)
    Public Noc2 As Integer          '[-] Number paralle Cyclones
    Public db2 As Double            '[m] Diameter cyclone body
    Public inh2 As Double           '[m] inlet hoogte
    Public inb2 As Double           '[m] inlet breedte
    Public dout2 As Double          '[m] diameter zakbuis
    Public inv2 As Double           '[m/s] Inlet velocity cyclone
    Public outv2 As Double          '[m/s] Outlet velocity cyclone
    Public Kstokes2 As Double       'Stokes of the particle

    Public m2 As Double             'm factor loss curve d< dia critical
    Public stage2() As GvG_Calc_struct   'tbv calculatie stage #2
    Public Dmin2 As Double          'Smallest particle in calculation
    Public Dmax2 As Double          'Biggest particle in calculation
End Structure

'Variables used by GvG in calculation
<Serializable()> Public Structure GvG_Calc_struct
    Public dia As Double            '[mu] Particle diameter 
    Public d_ave As Double          '[mu] Average diameter 
    Public d_ave_K As Double        '[-] Average diam/K_stokes 
    Public i_grp As Double          'Particle Groepnummer (stage 2= stage 1)
    Public i_d1 As Double           '[mu] particle diameter lower Class
    Public i_d2 As Double           '[mu] particle diameter upper Class
    Public i_p1 As Double           '[%] User Input percentage 1
    Public i_p2 As Double           '[%] User Input percentage 2
    Public i_k As Double            '[-] Parameter k
    Public i_m As Double            '[-] Parameter m
    Public psd_dif As Double        '[%] PSD diff
    Public psd_cum As Double        '[-] interpolatie psd cummulatief
    Public psd_cum_pro As Double    '[%] PSD cummulatief in [procent] for chart1
    Public loss_abs As Double       '[g/Am3] loss abs
    Public loss_abs_C As Double     '[g/Am3] loss abs compensated
    Public loss_overall As Double   'Overall Corrected
    Public loss_overall_C As Double 'Overall loss Corrected
    Public catch_chart As Double    '[%] for chart
End Structure

Public Class Form1
    Public _cyl1_dim(20) As Double          'Cyclone stage #1 dimensions
    Public _cyl2_dim(20) As Double          'Cyclone stage #2 dimensions
    Public _istep As Double                 '[mu] Particle size calculation step
    Public _cees(20) As Input_struct        '20 Case's data
    Dim k41 As Double                       'sum loss abs (for DataGridView1)

    'Type AC;Inlaatbreedte;Inlaathoogte;Inlaatlengte;Inlaat hartmaat;Inlaat afschuining;
    'Uitlaat keeldia inw.;Uitlaat flensdiameter inw.;Lengte insteekpijp inw.;
    'Lengte romp + conus;Lengte romp;Lengte conus;Dia_conus / 3P-pijp;Lengte 3P-pijp;Lengte 3P-conus;Kleine dia 3P-conus",

    ReadOnly cyl_dimensions() As String = {
    "AC-300;0.34;0.77;0.6;0.63;0.3;0.68;0.68;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-350;0.32;0.7;0.6;0.617;0.3;0.63;0.63;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-435;0.282;0.64;0.6;0.6;0.3;0.56;0.56;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-550;0.25;0.57;0.6;0.58;0.3;0.45;0.56;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-750;0.216;0.486;0.6;0.57;0.3;0.365;0.56;0.892;3.36;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-850;0.203;0.457;0.6;0.564;0.3;0.307;0.428;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-850+afz;0.203;0.457;0.6;0.564;0.3;0.307;0.428;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-1850;0.136;0.31;0.6;0.53;0.3;0.15;0.25;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-1850+afz;0.136;0.31;0.6;0.53;0.3;0.15;0.25;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25"}

    'DSM polymer power DSM Geleen (cumulatief[%], particle diameter[mu])
    ReadOnly DSM_psd_example() As String = {
    "4.50;0.04",
    "5.50;0.27",
    "6.50;0.61",
    "7.50;1.06",
    "9.00;1.98",
    "11.0;3.67",
    "13.0;5.93",
    "15.5;9.50",
    "18.5;14.73",
    "21.5;20.71",
    "25.0;28.1",
    "30.0;38.59",
    "37.5;52.93",
    "45.0;64.66",
    "52.5;73.66",
    "62.5;82.16",
    "75.0;88.86",
    "90.0;93.38",
    "105;96.05",
    "125;97.86",
    "150;98.92",
    "180;99.58",
    "215;99.95",
    "255;100",
    "305;100",
    "365;100",
    "435;100",
    "515;100",
    "615;100",
    "735;100",
    "875;100"}


    'Nieuwe reken methode, verdeling volgens Weibull verdeling
    'm1,k1,a1 als d < d_krit
    'm2,k2,a2 als d > d_krit
    'type; d/krit; m1; k1; a1; m2; k2; a2; drukcoef air;drukcoef dust
    ReadOnly rekenlijnen() As String = {
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
    ReadOnly dirpath_Eng As String = "N:\Engineering\VBasic\Cyclone_sizing_input\"
    ReadOnly dirpath_Rap As String = "N:\Engineering\VBasic\Cyclone_rapport_copy\"
    ReadOnly dirpath_tmp As String = "C:\Tmp\"
    ReadOnly ProcID As Integer = Process.GetCurrentProcess.Id
    ReadOnly dirpath_Temp As String = "C:\Temp\" & ProcID.ToString

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

        DataGridView1.ColumnCount = 10
        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add(23)
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells

        'Initialize the arrays in the struct
        For i = 0 To _cees.Length - 1
            _cees(i).case_name = ""
            ReDim _cees(i).dia_big(11)          'Initialize
            ReDim _cees(i).class_load(11)       'Initialize
            ReDim _cees(i).stage1(150)          'Initialize
            ReDim _cees(i).stage2(150)          'Initialize
        Next

        '------ allowed users with hard disc id's -----
        user_list.Add("user")
        hard_disk_list.Add("058F63646471")          'Privee PC, graslaan25

        user_list.Add("GerritP")
        hard_disk_list.Add("S2R6NX0H740154H")       'VTK PC, GP

        user_list.Add("GerritP")
        hard_disk_list.Add("0008_0D02_003E_0FBB.")  'VTK laptop, GP

        user_list.Add("GP")
        hard_disk_list.Add("S28ZNXAG521979")        'VTK laptop, GP privee

        user_list.Add("FredKo")
        hard_disk_list.Add("JR10006P02Y6EE")        'VTK laptop, FKo

        user_list.Add("JanK")
        hard_disk_list.Add("0025_38B4_71B4_88FC.")  'VTK laptop, Jank

        user_list.Add("JeroenA")
        hard_disk_list.Add("171095402070")          'VTK desktop, Jeroen

        user_list.Add("JeroenA")
        hard_disk_list.Add("170228801578")          'VTK laptop, Jeroen disk 1
        hard_disk_list.Add("MCDBM1M4F3QRBEH6")      'VTK laptop, Jeroen disk 2
        hard_disk_list.Add("0025_388A_81BB_14B5.")  'Zweet kamer, Jeroen 

        user_list.Add("Peterdw")
        hard_disk_list.Add("134309552747")          'VTK PC, Peter de Wild

        user_list.Add("Abi")
        hard_disk_list.Add("174741803447")          'VTK Laptop, Ab van Iterson <a.van.Iterson@vtk.nl>

        nu = Now()
        nu2 = CDate("2020-04-01 00:00:00")
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

        TextBox148.Text = "Het d50 getal geeft de diameter aan waarbij 50% gevangen wordt en 50% verloren gaat." & vbCrLf
        TextBox148.Text &= " " & vbCrLf
        TextBox148.Text &= "In de cirkel is voor iedere cycloon de verhouding van het d50 getal tov" & vbCrLf
        TextBox148.Text &= "het d50 getal voor AC435 cycloon bij het debiet Qv aangegeven." & vbCrLf
        TextBox148.Text &= "Het oppervlak van het instroom kanaal is voor alle modellen (bij gelijk debiet) ongeveer gelijk," & vbCrLf
        TextBox148.Text &= "de cycloon body is echter geheel verschillend." & vbCrLf
        TextBox148.Text &= " " & vbCrLf
        TextBox148.Text &= "Hoe groter de cycloon hoe beter hij vangt." & vbCrLf
        TextBox148.Text &= "Kenmerk voor hoge efficiency is dat de cycloon groter is dan de standaard cycloon" & vbCrLf
        TextBox148.Text &= "Een betere vangst zou aan een langere verblijftijd toegeschreven kunnen worden." & vbCrLf
        TextBox148.Text &= " " & vbCrLf
        TextBox148.Text &= "Een goede combinatie blijkt voorvangen met een AC435 gevolgd door twee stuks AC850." & vbCrLf
        TextBox148.Text &= " " & vbCrLf
        TextBox148.Text &= "Note 1) Spray-dried dairy products are fragile and will break up in a cyclone into several smaller parts," & vbCrLf
        TextBox148.Text &= "in this case the inlet speed must be reduced to 14-16 m/s(also depends on moisture content)." & vbCrLf
        TextBox148.Text &= "Patato starch does the opposite, less particles come out than go in. " & vbCrLf
        TextBox148.Text &= " " & vbCrLf



        TextBox20.Text = "All AA cyclones have a diameter of 300mm" & vbCrLf
        TextBox20.Text &= "Load above 5 gr/m3 is considered a high load" & vbCrLf
        TextBox20.Text &= "Cyclones can not choke" & vbCrLf

        TextBox47.Text = "Applications" & vbCrLf
        TextBox47.Text &= "Fly catcher Venezuela before a gasturbine" & vbCrLf
        TextBox47.Text &= "Spark catcher (metalic particles)" & vbCrLf
        TextBox47.Text &= "Droplet catcher" & vbCrLf
        TextBox47.Text &= "Patato Starch" & vbCrLf

        TextBox126.Text = "Calculation method" & vbCrLf
        TextBox126.Text &= "De cycloon instroom wordt verdeeld in 110 korrel fracties gebaseerd op de diameter " & vbCrLf
        TextBox126.Text &= "van de kleinste naar de grootste deeltje" & vbCrLf
        TextBox126.Text &= "Start met het kleinste klasse van trap 2" & vbCrLf
        TextBox126.Text &= "Iedere fractie heeft een verlies berekening gebaseerd op Weibull formule" & vbCrLf
        TextBox126.Text &= "Bepaal Stokes voor de het gemiddelde deeltje en zoek m, k, a voor de gebruikte cycloon." & vbCrLf
        TextBox126.Text &= "Opgelet m,k,a zijn verschillen voor deeltje groter en kleiner dan de kritische diameter" & vbCrLf
        TextBox126.Text &= "De kritische diameter is gegeven door het cycloon model (AC300, AC350 enz)" & vbCrLf
        TextBox126.Text &= "" & vbCrLf

        TextBox126.Text &= "De klant levert een korrelgroote verdeling zeg 10 fracties(psd= particle size distribution)" & vbCrLf
        TextBox126.Text &= "Tussen 2 punten van de korrel verdeling word met de Rosin Rammler formule een cumulatieve korrelverdeling getrokken" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "Bepaal voor iedere van de 110 korrelfracties" & vbCrLf
        TextBox126.Text &= "diameter -min[mu], -max[mu], -gem[mu], Stokes en voor de gekozen cycloon m, k, a" & vbCrLf
        TextBox126.Text &= "Groep nummer (in dit geval 1..10), d1(kleinste diameter vd groep), d2 (grootste)," & vbCrLf
        TextBox126.Text &= "p1(cumulatief gewicht groep vd ondergrens (0.99990-0.001), p2(cum. gew. bovengrens)" & vbCrLf
        TextBox126.Text &= "Bepaal het verlies getal 0-1 (0 voor de grote deeltjes en 1 voor de hele kleine deeltjes) " & vbCrLf
        TextBox126.Text &= "Verwerk correctie voor de stofbelading (hoog > 1-5 gr/m3) en insteekpijp" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "We kunnen nu het verlies per fractie berekenen loss*psd_diff= loss_abs" & vbCrLf
        TextBox126.Text &= "psd_dif gesommeerd over de 110 fracties moet de 1.0 (100%) naderen." & vbCrLf
        TextBox126.Text &= "Emissie= Total_loss * inlaat_stof_belasting" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "Opmerking 1) deeltjes < 0.5-0.7 mu kunnen niet gevangen worden ivm fysische mechanismen groter dan de centrifugaal kracht."

        TextBox145.Text = "Particles may breakup into smaller ones inside a cyclone" & vbCrLf
        TextBox145.Text &= "See project P10.1070"

        TextBox147.Text = "Cyclone is a excellent Spark arrestor" & vbCrLf
        TextBox147.Text &= "Use in front of filter or silo" & vbCrLf
        TextBox147.Text &= "Also used in suction of gas-turbines to catch flies" & vbCrLf

        TextBox149.Text = "Gluten separation with a Cyclone" & vbCrLf
        TextBox149.Text &= "Test have show that the particles are getting electrically charged and stick to the vessel wall" & vbCrLf

        TextBox152.Text = "Particle Density and Bulk (Volumetric) Density" & vbCrLf
        TextBox152.Text &= "Rule of thumb for foodstuffs, Particle Density= 2 x Bulk density" & vbCrLf

        TextBox153.Text = "Spray dried product are fragile" & vbCrLf
        TextBox153.Text &= "Inlet speed cyclone is 16 m/s" & vbCrLf
        TextBox153.Text &= "Spray dryer outlet pressure range is between 0 and -5 mbar" & vbCrLf
        Calc_sequence()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles button1.Click, TabPage1.Enter, numericUpDown3.ValueChanged, numericUpDown2.ValueChanged, numericUpDown14.ValueChanged, NumericUpDown1.ValueChanged, numericUpDown5.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown18.ValueChanged, ComboBox1.SelectedIndexChanged, numericUpDown9.ValueChanged, numericUpDown8.ValueChanged, numericUpDown7.ValueChanged, numericUpDown6.ValueChanged, numericUpDown12.ValueChanged, numericUpDown11.ValueChanged, numericUpDown10.ValueChanged, numericUpDown13.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown27.ValueChanged, NumericUpDown26.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown24.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown34.ValueChanged, NumericUpDown33.ValueChanged, ComboBox2.SelectedIndexChanged, NumericUpDown40.ValueChanged, NumericUpDown39.ValueChanged, NumericUpDown38.ValueChanged, NumericUpDown37.ValueChanged, NumericUpDown36.ValueChanged, NumericUpDown35.ValueChanged, NumericUpDown43.ValueChanged, NumericUpDown22.ValueChanged, CheckBox3.CheckedChanged, CheckBox2.CheckedChanged
        Calc_sequence()
    End Sub

    Private Sub Get_input_calc_1(ks As Integer)
        Dim db1 As Double           'Body diameter stage #1
        Dim db2 As Double           'Body diameter stage #2
        Dim words() As String

        '===Input parameter ===
        Dim tot_kgh As Double       'Dust inlet per hour totaal 
        Dim ro_solid As Double      'Density [kg/hr]
        Dim visco As Double         'Visco in Centi Poise
        Dim ratio As Double         'ratio  kg/Nm3 and kg/Am3

        '==== data ======
        Dim wc_dust1, wc_dust2 As Double    'weerstand_coef_air
        Dim wc_air1, wc_air2 As Double      'weerstand_coef_air

        '==== results ===
        Dim kgh As Double               'Dust inlet per hour/cyclone 
        Dim kgs As Double               'Dust inlet per second/cyclone
        Dim kortezijde, langezijde As Double    'tbv atan()
        Dim halfconeapex1 As Double     'Conus hoek
        Dim halfconeapex2 As Double     'Conus hoek

        '==== stage 1 ====

        '------ dust load NORMAL conditions ----
        _cees(ks).Ro_gas1_Am3 = numericUpDown3.Value                    '[kg/Am3]
        _cees(ks).Ro_gas1_Nm3 = Calc_Normal_density(_cees(ks).Ro_gas1_Am3, _cees(ks).p1_abs, _cees(ks).Temp)

        '--------- ratio  Nm2 and Am3 ---------
        ratio = _cees(ks).Ro_gas1_Nm3 / _cees(ks).Ro_gas1_Am3
        _cees(ks).dust1_Am3 = NumericUpDown4.Value                    'gram/Am3
        _cees(ks).dust1_Nm3 = _cees(ks).dust1_Am3 * ratio               'gram/Nm3

        '--------- present -------------
        TextBox132.Text = _cees(ks).dust1_Nm3.ToString("F2")          'gram/Nm3
        TextBox129.Text = _cees(ks).Ro_gas1_Nm3.ToString("F3")        'kg/Nm3 inlet gas

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
            _cees(ks).p1_abs = NumericUpDown19.Value * 100 + 101325  'Pressure [mbar]

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

            '------ determine high/low dust load ------------
            CheckBox2.Checked = CBool(IIf(_cees(ks).dust1_Am3 > 20, True, False))
            _cees(ks).FlowT = NumericUpDown1.Value      '[m3/h] 
            _cees(ks).Flow1 = _cees(ks).FlowT / (3600 * _cees(ks).Noc1) '[Am3/s/cycloon]

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
            _cees(ks).dpgas1 = 0.5 * _cees(ks).Ro_gas1_Am3 * _cees(ks).inv1 ^ 2 * wc_air1      '[Pa]
            _cees(ks).dpdust1 = 0.5 * _cees(ks).Ro_gas1_Am3 * _cees(ks).inv1 ^ 2 * wc_dust1    '[Pa]
            _cees(ks).p2_abs = _cees(ks).p1_abs - _cees(ks).dpgas1              '[P_abs inlet]


            '=========== Stage #2 ==============
            _cees(ks).Ro_gas2_Am3 = _cees(ks).Ro_gas1_Am3 * _cees(ks).p2_abs / _cees(ks).p1_abs  '[kg/Am3]
            _cees(ks).Ro_gas2_Nm3 = Calc_Normal_density(_cees(ks).Ro_gas2_Am3, _cees(ks).p2_abs, _cees(ks).Temp)

            _cees(ks).Flow2 = _cees(ks).FlowT / (3600 * _cees(ks).Noc2)         '[Am3/s/cycloon]

            '------ compensate the flow for the pressure loss over stage 1 ---
            _cees(ks).Flow2 *= _cees(ks).p1_abs / _cees(ks).p2_abs

            '---- Compensate for the Speed for the pressure loss in stage #1 ----
            _cees(ks).inv2 = _cees(ks).Flow2 / (_cees(ks).inb2 * _cees(ks).inh2)    '[m/s]
            _cees(ks).outv2 = _cees(ks).Flow2 / ((PI / 4) * _cees(ks).dout2 ^ 2)    '[m/s]

            '----------- Pressure loss cyclone stage #2----------------------
            If ComboBox2.SelectedIndex > -1 Then
                words = rekenlijnen(ComboBox2.SelectedIndex).Split(CType(";", Char()))
                wc_air2 = CDbl(words(8))
                wc_dust2 = CDbl(words(9))
            End If
            _cees(ks).dpgas2 = 0.5 * _cees(ks).Ro_gas2_Am3 * _cees(ks).inv2 ^ 2 * wc_air2
            _cees(ks).dpdust2 = 0.5 * _cees(ks).Ro_gas2_Am3 * _cees(ks).inv2 ^ 2 * wc_dust2

            '----------- 1/2 cone apex #1-----------
            kortezijde = (db1 - _cyl1_dim(12) * db1) * 0.5
            langezijde = _cyl1_dim(11) * db1
            halfconeapex1 = Atan(kortezijde / langezijde)           '[rad]
            halfconeapex1 = halfconeapex1 / (PI / 2) * 90           '[degree]

            '----------- 1/2 cone apex #2-----------
            kortezijde = (db2 - _cyl1_dim(12) * db2) * 0.5
            langezijde = _cyl1_dim(11) * db2
            halfconeapex2 = Atan(kortezijde / langezijde)           '[rad] 
            halfconeapex2 = halfconeapex2 / (PI / 2) * 90           '[degree]

            '----------- stof belasting ------------
            kgs = _cees(ks).Flow1 * _cees(ks).dust1_Am3 / 1000  '[kg/s/cycloon]
            kgh = kgs * 3600                                    '[kg/h/cycloon]
            tot_kgh = kgh * _cees(ks).Noc1                      '[g/Am3] Dust inlet 

            '----------- K_stokes-----------------------------------
            _cees(ks).Kstokes1 = Sqrt(db1 * 2000 * visco * 16 / (ro_solid * 0.0181 * _cees(ks).inv1))
            _cees(ks).Kstokes2 = Sqrt(db2 * 2000 * visco * 16 / (ro_solid * 0.0181 * _cees(ks).inv2))

            '----------- presenteren ----------------------------------
            TextBox128.Text = _cees(ks).Ro_gas2_Nm3.ToString("F3")      '[kg/Nm3] density
            TextBox36.Text = (_cees(ks).FlowT / 3600).ToString("F4")    '[m3/s] flow

            '----------- presenteren afmetingen in [mm] ----------------------------
            TextBox1.Text = (_cees(ks).inh1).ToString("F3")          'inlaat breedte
            TextBox2.Text = (_cees(ks).inb1).ToString("F3")          'Inlaat hoogte
            TextBox3.Text = (_cyl1_dim(3) * db1).ToString("F3")      'Inlaat lengte
            TextBox4.Text = (_cyl1_dim(4) * db1).ToString("F3")      'Inlaat hartmaat
            TextBox5.Text = (_cyl1_dim(5) * db1).ToString("F3")      'Inlaat afschuining
            TextBox6.Text = (_cyl1_dim(6) * db1).ToString("F3")      'Uitlaat keeldia inw.
            TextBox7.Text = (_cyl1_dim(7) * db1).ToString("F3")      'Uitlaat flensdiameter inw.
            TextBox8.Text = (_cyl1_dim(8) * db1).ToString("F3")      'Lengte insteekpijp inw.
            TextBox9.Text = (_cyl1_dim(9) * db1).ToString("F3")      'Lengte romp + conus
            TextBox10.Text = (_cyl1_dim(10) * db1).ToString("F3")    'Lengte romp
            TextBox11.Text = (_cyl1_dim(11) * db1).ToString("F3")    'Lengte çonus
            TextBox12.Text = (_cyl1_dim(12) * db1).ToString("F3")    'Dia_conus / 3P-pijp
            TextBox13.Text = (_cyl1_dim(13) * db1).ToString("F3")    'Lengte 3P-pijp
            TextBox14.Text = (_cyl1_dim(14) * db1).ToString("F3")    'Lengte 3P conus
            TextBox15.Text = (_cyl1_dim(15) * db1).ToString("F3")    'Kleine dia 3P-conus

            TextBox84.Text = (_cees(ks).inh2).ToString("F3")         'inlaat breedte
            TextBox85.Text = (_cees(ks).inb2).ToString("F3")         'Inlaat hoogte
            TextBox86.Text = (_cyl1_dim(3) * db2).ToString("F3")     'Inlaat lengte
            TextBox87.Text = (_cyl1_dim(4) * db2).ToString("F3")     'Inlaat hartmaat
            TextBox88.Text = (_cyl1_dim(5) * db2).ToString("F3")     'Inlaat afschuining
            TextBox89.Text = (_cyl1_dim(6) * db2).ToString("F3")     'Uitlaat keeldia inw.
            TextBox90.Text = (_cyl1_dim(7) * db2).ToString("F3")     'Uitlaat flensdiameter inw.
            TextBox91.Text = (_cyl1_dim(8) * db2).ToString("F3")     'Lengte insteekpijp inw.
            TextBox92.Text = (_cyl1_dim(9) * db2).ToString("F3")     'Lengte romp + conus
            TextBox93.Text = (_cyl1_dim(10) * db2).ToString("F3")    'Lengte romp
            TextBox94.Text = (_cyl1_dim(11) * db2).ToString("F3")    'Lengte çonus
            TextBox95.Text = (_cyl1_dim(12) * db2).ToString("F3")    'Dia_conus / 3P-pijp
            TextBox96.Text = (_cyl1_dim(13) * db2).ToString("F3")    'Lengte 3P-pijp
            TextBox97.Text = (_cyl1_dim(14) * db2).ToString("F3")    'Lengte 3P conus
            TextBox98.Text = (_cyl1_dim(15) * db2).ToString("F3")    'Kleine dia 3P-conus
            TextBox150.Text = halfconeapex1.ToString("F1")           '1/2 cone apex
            TextBox151.Text = halfconeapex2.ToString("F1")           '1/2 cone apex

            TextBox113.Text = (_cees(ks).Flow1 * 3600).ToString("0")    '[Am3/s] Cycloone Flow
            TextBox112.Text = (_cees(ks).Flow2 * 3600).ToString("0")    '[Am3/s] Cycloone Flow

            TextBox16.Text = _cees(ks).inv1.ToString("0.0")             'inlaat snelheid
            TextBox80.Text = _cees(ks).inv2.ToString("0.0")             'inlaat snelheid

            TextBox17.Text = _cees(ks).dpgas1.ToString("0")             '[Pa] Pressure loss inlet-gas
            TextBox79.Text = _cees(ks).dpgas2.ToString("0")             '[Pa] Pressure loss inlet-gas

            TextBox48.Text = _cees(ks).dpdust1.ToString("0")            '[Pa] Pressure loss inlet-dust
            TextBox76.Text = _cees(ks).dpdust2.ToString("0")            '[Pa] Pressure loss inlet-dust

            TextBox22.Text = _cees(ks).outv1.ToString("0.0")            'uitlaat snelheid
            TextBox77.Text = _cees(ks).outv2.ToString("0.0")            'uitlaat snelheid

            TextBox23.Text = _cees(ks).Kstokes1.ToString("F4")          'Stokes waarde stage#1
            TextBox78.Text = _cees(ks).Kstokes2.ToString("F4")          'Stokes waarde stage#2

            TextBox37.Text = _cees(ks).db1.ToString                     'Cycloone diameter stage#1
            TextBox74.Text = _cees(ks).db2.ToString                     'Cycloone diameter stage#2

            TextBox38.Text = CType(ComboBox1.SelectedItem, String)      'Cycloon type
            TextBox73.Text = CType(ComboBox2.SelectedItem, String)      'Cycloon type

            '---------- Pressure abs --------------
            TextBox131.Text = _cees(ks).p1_abs.ToString("F0")           '[Pa abs]
            TextBox130.Text = _cees(ks).p2_abs.ToString("F0")           '[Pa abs]

            '---------- Density --------------
            TextBox75.Text = _cees(ks).Ro_gas1_Am3.ToString("F3")       '[kg/Am3]
            TextBox19.Text = _cees(ks).Ro_gas2_Am3.ToString("F3")       '[kg/Am3]

            '---------- Check inlet speed [m/s] stage #1---------------
            If _cees(ks).inv1 < 10 Or _cees(ks).inv1 > 30 Then
                TextBox16.BackColor = Color.Red
            Else
                TextBox16.BackColor = Color.LightGreen
            End If

            '---------- Check inlet speed stage #2---------------
            If _cees(ks).inv2 < 10 Or _cees(ks).inv2 > 30 Then
                TextBox80.BackColor = Color.Red
            Else
                TextBox80.BackColor = Color.LightGreen
            End If

            '---------- Check dp [pa] stage #1---------------
            If _cees(ks).dpgas1 > 3000 Then
                TextBox17.BackColor = Color.Red
            Else
                TextBox17.BackColor = Color.LightGreen
            End If

            '---------- Check dp stage #2---------------
            If _cees(ks).dpgas2 > 3000 Then
                TextBox79.BackColor = Color.Red
            Else
                TextBox79.BackColor = Color.LightGreen
            End If

            '--------- Get Inlet korrel-groep data ----------
            'Save data of screen into the _cees array
            Fill_array_from_screen(CInt(NumericUpDown30.Value))

            TextBox39.Text = kgh.ToString("F0")                 'Stof inlet [kg/(h.cyclone)]
            TextBox40.Text = tot_kgh.ToString("F0")             'Dust inlet [g/Am3] 
            TextBox71.Text = _cees(ks).dust1_Am3.ToString("F1") 'Dust inlet [g/Am3]
        End If
    End Sub
    Private Sub Present_Datagridview1(ks As Integer)

        '--------- HeaderText --------------------
        DataGridView1.Columns(0).HeaderText = "Dia class"
        DataGridView1.Columns(1).HeaderText = "Feed psd cum"
        DataGridView1.Columns(2).HeaderText = "Feed psd diff"
        DataGridView1.Columns(3).HeaderText = "Loss [%] of feed"
        DataGridView1.Columns(4).HeaderText = "Loss abs [%]"
        DataGridView1.Columns(5).HeaderText = "Loss psd cum"
        DataGridView1.Columns(6).HeaderText = "Catch abs"
        DataGridView1.Columns(7).HeaderText = "Catch psd cum"
        DataGridView1.Columns(8).HeaderText = "Grade class eff."

        Calc_Datagridview1(ks)
        '=========  sum is required for column 5 calculation ===============
        k41 = 0
        For h = 0 To 22
            k41 += CDbl(DataGridView1.Rows(h).Cells(4).Value)    'tot_catch_abs[%]
        Next
        '===================================================================

        DataGridView1.AutoResizeColumns()
    End Sub
    Private Sub Calc_Datagridview1(ks As Integer)

        '==== stage #1 + stage #2 ====
        Dim h18, h19 As Double
        Dim j18, i18 As Double
        Dim l18, k19 As Double
        Dim k18 As Double
        Dim m18, n17_oud, n18 As Double
        Dim tot_catch_abs As Double
        Dim o18 As Double
        Dim tt As Double

        For h = 0 To 22
            DataGridView1.Rows(h).Cells(0).Value = _cees(ks).stage1(h * 5).d_ave.ToString("F3") 'diameter
            DataGridView1.Rows(h).Cells(1).Value = _cees(ks).stage1(h * 5).psd_cum_pro.ToString("F3") 'feed psd cum

            If h > 0 Then
                h18 = CDbl(DataGridView1.Rows(h - 1).Cells(1).Value)
            Else
                h18 = 100
            End If

            h19 = CDbl(DataGridView1.Rows(h).Cells(1).Value)   'feed psd cum
            DataGridView1.Rows(h).Cells(2).Value = (h18 - h19).ToString("F3")   'feed psd diff

            '========= (column 3) loss ===============
            If CheckBox2.Checked Then
                DataGridView1.Rows(h).Cells(3).Value = (_cees(ks).stage1(h * 5).loss_overall_C * 100).ToString("F3")
            Else
                DataGridView1.Rows(h).Cells(3).Value = (_cees(ks).stage1(h * 5).loss_overall * 100).ToString("F3")
            End If

            i18 = CDbl(DataGridView1.Rows(h).Cells(2).Value) 'feed psd diff
            j18 = CDbl(DataGridView1.Rows(h).Cells(3).Value) 'loss % Of feed
            DataGridView1.Rows(h).Cells(4).Value = (i18 * j18 / 100).ToString("F3") 'Loss abs [%]

            '=========  (column 4) Loss abs [%] ===============
            If h > 0 Then
                l18 = CDbl(DataGridView1.Rows(h - 1).Cells(5).Value)
            Else
                l18 = 100
            End If
            k19 = CDbl(DataGridView1.Rows(h).Cells(4).Value)   'Loss abs [%]

            '=========  (column 5) Loss psd cum ===============
            If h > 0 Then
                l18 = CDbl(DataGridView1.Rows(h - 1).Cells(5).Value)
            Else
                l18 = 100
            End If
            tt = (l18 - 100 * k19 / k41)
            If tt < 0 Then tt = 0           'Prevent negative numbers
            DataGridView1.Rows(h).Cells(5).Value = tt.ToString("F3")

            '============= (column 6) Loss abs [%] ===================
            k18 = CDbl(DataGridView1.Rows(h).Cells(4).Value)   'Loss abs [%]
            m18 = (i18 - k18)
            DataGridView1.Rows(h).Cells(6).Value = m18.ToString("F3") 'Catch abs

            '=============  (column 7) Catch psd cum  ===================
            Double.TryParse(TextBox59.Text, tot_catch_abs)      'tot_catch_abs[%]
            If h > 0 Then
                n17_oud = CDbl(DataGridView1.Rows(h - 1).Cells(7).Value)
                n18 = n17_oud - m18 / (tot_catch_abs / 100)
            Else
                n18 = 100
            End If
            n18 = CDbl(IIf(n18 < 0, 0, n18))        'prevent silly results
            DataGridView1.Rows(h).Cells(7).Value = n18.ToString("F3") 'Catch psd cum

            '=========  (column 8) Efficiency ===============
            o18 = 100 - j18
            DataGridView1.Rows(h).Cells(8).Value = o18.ToString("F3")           'Grade eff.
        Next h
    End Sub

    Private Sub Calc_part_dia_loss(ks As Integer)

        '---------- Calc particle diameter with x% loss ---
        '---------- present stage #1 -------
        TextBox42.Text = Calc_dia_particle(1.0, _cees(ks).Kstokes1, 1).ToString("F2")     '[mu] @ 100% loss
        TextBox26.Text = Calc_dia_particle(0.95, _cees(ks).Kstokes1, 1).ToString("F2")    '[mu] @  95% lost
        TextBox31.Text = Calc_dia_particle(0.9, _cees(ks).Kstokes1, 1).ToString("F2")     '[mu] @  90% lost
        TextBox32.Text = Calc_dia_particle(0.5, _cees(ks).Kstokes1, 1).ToString("F2")     '[mu] @  50% lost
        TextBox33.Text = Calc_dia_particle(0.1, _cees(ks).Kstokes1, 1).ToString("F2")     '[mu] @  10% lost
        TextBox41.Text = Calc_dia_particle(0.05, _cees(ks).Kstokes1, 1).ToString("F2")    '[mu] @   5% lost

        '---------- present stage #2 -------
        ' MessageBox.Show(_cees(ks).Kstokes2.ToString)
        TextBox102.Text = Calc_dia_particle(1.0, _cees(ks).Kstokes2, 2).ToString("F2")     '[mu] @  100% lost
        TextBox103.Text = Calc_dia_particle(0.95, _cees(ks).Kstokes2, 2).ToString("F2")    '[mu] @  95% lost
        TextBox104.Text = Calc_dia_particle(0.9, _cees(ks).Kstokes2, 2).ToString("F2")     '[mu] @  90% lost
        TextBox105.Text = Calc_dia_particle(0.5, _cees(ks).Kstokes2, 2).ToString("F2")     '[mu] @  50% lost
        TextBox106.Text = Calc_dia_particle(0.1, _cees(ks).Kstokes2, 2).ToString("F2")     '[mu] @  10% lost
        TextBox107.Text = Calc_dia_particle(0.05, _cees(ks).Kstokes2, 2).ToString("F2")    '[mu] @   5% lost
    End Sub

    Private Sub Fill_array_from_screen(c_nr As Integer)

        TextBox24.Text &= "line 660, Fill array from screen, nr " & c_nr.ToString & vbCrLf

        _cees(c_nr).Quote_no = TextBox28.Text                  'Quote number
        _cees(c_nr).Tag_no = TextBox29.Text                    'The Tag number
        _cees(c_nr).case_name = TextBox53.Text                  'The case name

        _cees(c_nr).FlowT = NumericUpDown1.Value                'Air flow [Am3/hr]
        _cees(c_nr).dust1_Am3 = NumericUpDown4.Value            'Dust inlet [g/Am3] 
        _cees(c_nr).Ct1 = ComboBox1.SelectedIndex               'Cyclone type Stage #1
        _cees(c_nr).Ct2 = ComboBox2.SelectedIndex               'Cyclone type Stage #2
        _cees(c_nr).Noc1 = CInt(NumericUpDown20.Value)          'Cyclone in parallel
        _cees(c_nr).Noc2 = CInt(NumericUpDown33.Value)          'Cyclone in parallel
        _cees(c_nr).db1 = numericUpDown13.Value                 'Diameter cyclone Stage #1
        _cees(c_nr).db2 = NumericUpDown34.Value                 'Diameter cyclone Stage #2
        _cees(c_nr).ro_gas = numericUpDown3.Value               'Density [kg/hr]
        _cees(c_nr).ro_solid = numericUpDown2.Value             'Density [kg/hr]
        _cees(c_nr).visco = numericUpDown14.Value               'Visco in [Centi Poise]
        _cees(c_nr).Temp = NumericUpDown18.Value                'Temperature [c]
        _cees(c_nr).p1_abs = 101325 + (NumericUpDown19.Value * 100)         'Pressure [Pa abs]


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
        _cees(c_nr).dia_big(10) = NumericUpDown37.Value  '80


        'Percentale van de inlaat stof belasting [%]
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
        NumericUpDown40.BackColor = CType(IIf(NumericUpDown40.Value > NumericUpDown48.Value, Color.LightGreen, Color.Red), Color)
        NumericUpDown48.BackColor = CType(IIf(NumericUpDown48.Value >= 0, Color.LightGreen, Color.Red), Color)
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
            Else
                verlies = 1.0        '100% loss (very small particle)
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
            If (stage = 1) Then                         'Stage #1 cyclone
                cor1 = NumericUpDown22.Value            'Correctie insteek pijp stage #1
                Double.TryParse(TextBox55.Text, cor2)   'Hoge stof belasting correctie acc VT-UK

            Else                                        'Stage #2 cyclone
                cor1 = NumericUpDown43.Value            'Correctie insteek pijp stage #2
                Double.TryParse(TextBox67.Text, cor2)   'Hoge stof belasting correctie acc VT-UK
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
        If (stage = 1) Then                         'Cyclone 1# stage
            cor1 = NumericUpDown22.Value            'Correctie insteek pijp stage #1
            Double.TryParse(TextBox55.Text, cor2)   'Hoge stof belasting correctie acc VT-UK
            words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
        Else                                        'Cyclone 2# stage
            cor1 = NumericUpDown43.Value            'Correctie insteek pijp stage #2
            Double.TryParse(TextBox67.Text, cor2)   'Hoge stof belasting correctie acc VT-UK
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
    Private Function Calc_dust_load_correction(dst As Double) As Double
        Dim f1, f2, f3, f4, f As Double
        '---- Dust load dimension is [kg/Am3] ---
        '---- Below 20 gram/Am3 NO correction ---

        f1 = 0.97833 + 2.918055 * dst - 39.3739 * dst ^ 2 + 472.0149 * dst ^ 3 - 769.586 * dst ^ 4
        f2 = -0.30338 + 21.91961 * dst - 73.5039 * dst ^ 2 + 112.485 * dst ^ 3 - 63.4408 * dst ^ 4
        f3 = 2.043212 + 0.725352 * dst - 0.2663 * dst ^ 2 + 0.04299 * dst ^ 3 - 0.00233 * dst ^ 4
        f4 = 2.853325 + 0.019026 * dst - 0.00036 * dst ^ 2 + 0.000003 * dst ^ 3 - 0.0000000065 * dst ^ 4

        Select Case dst
            Case < 0.02     '[20 gram/Am3]
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
        Return (f)
    End Function

    '---- According to VTK UK -----
    Private Sub Dust_load_correction(ks As Integer)
        Dim f_used As Double

        '============ stage 1 cyclone ==========
        f_used = Calc_dust_load_correction(_cees(ks).dust1_Am3 / 1000)
        TextBox55.Text = f_used.ToString("F3")

        '============ stage 2 cyclone ==========
        f_used = Calc_dust_load_correction(_cees(ks).dust2_Am3 / 1000)
        TextBox67.Text = f_used.ToString("F3")
    End Sub

    Public Sub Draw_chart1(ch As Chart)
        Dim ks As Integer
        Dim a, b As Double
        Dim eff1, eff2 As Double
        Dim loss1, loss2 As Double

        ch.Series.Clear()
        ch.ChartAreas.Clear()
        ch.Titles.Clear()
        ch.ChartAreas.Add("ChartArea0")

        For h = 0 To 6
            ch.Series.Add("Series" & h.ToString)
            ch.Series(h).ChartArea = "ChartArea0"
            ch.Series(h).ChartType = SeriesChartType.Line
            ch.Series(h).BorderWidth = 3
            ch.Series(h).IsVisibleInLegend = True
            'ch.Series(h).Label = "#VALY{F1}"
            ch.Series(h).IsValueShownAsLabel = False
        Next

        ch.Series(0).BorderWidth = 0    'Only markers NO line

        ch.Series(0).LegendText = "Input Stars stage #1 [%]   "
        ch.Series(1).LegendText = "Input Rosin Rammler #1 [%] "
        ch.Series(2).LegendText = "Loss stage #1 [%]          "
        ch.Series(3).LegendText = "Loss stage #2 [%]          "   '(staat niet in de GvG grafiek)
        ch.Series(4).LegendText = "Efficiency stage #1 [%]    "
        ch.Series(5).LegendText = "Efficiency stage #2 [%]    "
        ch.Series(6).LegendText = "Efficiency 1&2 [%]         "

        For h = 0 To 6
            If CheckBox13.Checked Then
                ch.Series(h).IsVisibleInLegend = True
            Else
                ch.Series(h).IsVisibleInLegend = False
            End If
        Next

        ch.Series(0).BorderDashStyle = ChartDashStyle.Solid
        ch.Series(1).BorderDashStyle = ChartDashStyle.Solid
        ch.Series(2).BorderDashStyle = ChartDashStyle.Solid
        ch.Series(3).BorderDashStyle = ChartDashStyle.Solid
        ch.Series(4).BorderDashStyle = ChartDashStyle.Dot
        ch.Series(5).BorderDashStyle = ChartDashStyle.Dot
        ch.Series(6).BorderDashStyle = ChartDashStyle.Dot

        ch.Series(0).IsValueShownAsLabel = CBool(IIf(CheckBox8.Checked, True, False))
        ch.Series(1).IsValueShownAsLabel = CBool(IIf(CheckBox9.Checked, True, False))
        ch.Series(2).IsValueShownAsLabel = CBool(IIf(CheckBox10.Checked, True, False))
        ch.Series(3).IsValueShownAsLabel = CBool(IIf(CheckBox1.Checked, True, False))
        ch.Series(4).IsValueShownAsLabel = CBool(IIf(CheckBox11.Checked, True, False))
        ch.Series(5).IsValueShownAsLabel = CBool(IIf(CheckBox12.Checked, True, False))
        ch.Series(6).IsValueShownAsLabel = CBool(IIf(CheckBox5.Checked, True, False))

        ch.ChartAreas("ChartArea0").AxisX.TitleFont = New Font("Arial", 11, System.Drawing.FontStyle.Bold)
        ch.ChartAreas("ChartArea0").AxisY.TitleFont = New Font("Arial", 11, System.Drawing.FontStyle.Bold)
        ch.Titles.Add("CALCULATED CUMULATIVE PARTICLE SIZE DISTRIBUTIONS")
        ch.Titles.Item(0).Font = New Font("Arial", 14, System.Drawing.FontStyle.Bold)

        ch.ChartAreas("ChartArea0").AxisX.Title = "particle dia [mu]"
        ch.ChartAreas("ChartArea0").AxisY.Title = "Exit loss [%]"
        ch.ChartAreas("ChartArea0").AxisY.Minimum = 0       'Loss
        ch.ChartAreas("ChartArea0").AxisY.Maximum = 100     'Loss
        ch.ChartAreas("ChartArea0").AxisY.Interval = 10     'Interval

        If CheckBox4.Checked Then
            ch.ChartAreas("ChartArea0").AxisX.MinorGrid.Enabled = True
            ch.ChartAreas("ChartArea0").AxisY.MinorGrid.Enabled = True
        End If

        If CheckBox7.Checked Then
            ch.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            ch.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True
        End If

        ch.ChartAreas("ChartArea0").AxisX.IsLogarithmic = True
        ch.ChartAreas("ChartArea0").AxisX.Minimum = 0.1     'Particle size
        ch.ChartAreas("ChartArea0").AxisX.Maximum = 100     'Particle size

        '----- now start plotting ------------------------
        ks = CInt(NumericUpDown30.Value)        'Case number

        '----------------------------- Plot Input stars-----------------------
        For input_cnt = 0 To 10
            'What is the star position
            a = _cees(ks).dia_big(input_cnt)              '[mu] Class upper particle diameter limit diameter
            b = _cees(ks).class_load(input_cnt) * 100     'Percentale van de inlaat stof belasting

            '--------------- plot-----------------
            ch.Series(0).Points.AddXY(a, b)
            ch.Series(0).Points(input_cnt).MarkerStyle = MarkerStyle.Star10
            ch.Series(0).Points(input_cnt).MarkerSize = 15
        Next

        '------ PSD Input Rosin Rammler stage #1 -------------
        For h = 0 To 110 Step 3   'Fill line chart
            a = _cees(ks).stage1(h).dia
            b = _cees(ks).stage1(h).psd_cum_pro
            ch.Series(1).Points.AddXY(a, b)
        Next h


        '------ Plot Stage #2 output-------------
        '------ Data from DataGridView2 -------------
        For h = 0 To DataGridView2.Rows.Count - 1                      'Fill line chart
            a = CDbl((DataGridView2.Rows(h).Cells(0).Value))           'Particle size
            eff1 = CDbl((DataGridView2.Rows(h).Cells(5).Value))        'Get data eff1
            eff2 = CDbl((DataGridView3.Rows(h).Cells(5).Value))        'Get data eff2

            '------- Plot efficiency --------
            ch.Series(4).Points.AddXY(a, eff1)                         'Plot eff #1
            If CheckBox6.Checked Then
                ch.Series(5).Points.AddXY(a, eff2)                     'Plot eff #2
            End If
        Next h

        '----- labels on chart ------------
        ch.Series(0).Points(5).Label = "Input"
        ch.Series(4).Points(30).Label = "Eff #1"
        If CheckBox6.Checked Then ch.Series(5).Points(45).Label = "Eff #2"

        '------ Data from DataGridView4 -------------
        For h = 0 To DataGridView4.Rows.Count - 1                       'Fill line chart
            a = CDbl((DataGridView4.Rows(h).Cells(0).Value))            'Particle size
            loss1 = CDbl((DataGridView4.Rows(h).Cells(5).Value))        'Get data Loss 1

            '------- Plot stage #1  --------
            ch.Series(2).Points.AddXY(a, loss1)                         'Plot Loss 1

            '------- Plot stage #2  --------
            If CheckBox6.Checked Then
                'loss2 = CDbl((DataGridView4.Rows(h).Cells(8).Value))   'get data Loss 2
                loss2 = CDbl((DataGridView3.Rows(h).Cells(14).Value))   'get data Loss 2

                b = CDbl((DataGridView4.Rows(h).Cells(9).Value))        'get data Eff 1&2 [%]
                ch.Series(3).Points.AddXY(a, loss2)                     'Plot Loss 2
                ch.Series(6).Points.AddXY(a, b)                         'Plot Eff 1&2 [%]
            End If
        Next h

    End Sub
    Public Sub Log_now(ks As Integer, line As Integer, r As String)
        'Dim verschil As Double

        'verschil = _cees(ks).stage1(line).psd_cum_pro - _cees(ks).stage2(line).psd_cum_pro
        'TextBox24.Text &= "Log line" & line.ToString & " " & r
        'TextBox24.Text &= " c.stage1(line).psd_dif=" & _cees(ks).stage1(line).psd_dif.ToString
        'TextBox24.Text &= "  c.stage2(line).psd_dif=" & _cees(ks).stage2(line).psd_dif.ToString
        'TextBox24.Text &= "  c.stage1(line).psd_cum_pro=" & _cees(ks).stage1(line).psd_cum_pro.ToString
        'TextBox24.Text &= "  c.stage2(line).psd_cum_pro=" & _cees(ks).stage2(line).psd_cum_pro.ToString
        'TextBox24.Text &= "  verschil=" & verschil.ToString & vbCrLf
    End Sub

    Private Sub Draw_chart2(ch As Chart)
        'Small chart on the first tab
        Dim s_points(50, 2) As Double
        Dim h As Integer
        Dim sdia As Integer
        Dim ks As Integer   'Case number

        ch.Series.Clear()
        ch.ChartAreas.Clear()
        ch.Titles.Clear()
        ch.ChartAreas.Add("ChartArea0")

        ch.Series.Add("Series1")
        ch.Series.Add("Series2")
        ch.Series(0).ChartArea = "ChartArea0"
        ch.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
        ch.Series(0).BorderWidth = 2
        ch.Series(0).IsVisibleInLegend = False

        ch.Series(1).ChartArea = "ChartArea0"
        ch.Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Line
        ch.Series(1).BorderWidth = 2
        ch.Series(1).IsVisibleInLegend = False

        ch.Titles.Add("Loss Curve Stage #1 and #2")
        ch.ChartAreas("ChartArea0").AxisX.Title = "Particle diameter [mu]"
        ch.ChartAreas("ChartArea0").AxisY.Title = "Cyclone loss [%]"
        ch.ChartAreas("ChartArea0").AxisY.Minimum = 0     'Loss
        ch.ChartAreas("ChartArea0").AxisY.Maximum = 100   'Loss
        ch.ChartAreas("ChartArea0").AxisX.Minimum = 0     'Particle size
        ch.ChartAreas("ChartArea0").AxisX.Maximum = 30    'Particle size

        '----- now calc chart points #1 and #2 --------------------------
        Integer.TryParse(TextBox42.Text, sdia)                  'dp(100) stage #1
        s_points(0, 0) = sdia                                   'Particle diameter [mu]
        s_points(0, 1) = 100                                    '100% loss
        s_points(0, 2) = 100                                    '100% loss
        ks = CInt(NumericUpDown30.Value)
        For h = 1 To 30                                          'Particle diameter [mu]
            s_points(h, 0) = h                                   'Particle diameter [mu]
            s_points(h, 1) = Calc_verlies(h, False, _cees(ks).Kstokes1, 1) * 100  'stage #1 Loss [%]
            s_points(h, 2) = Calc_verlies(h, False, _cees(ks).Kstokes2, 2) * 100  'stage #2 Loss [%]
        Next

        '------ now present-------------
        For h = 0 To 30 - 1   'Fill line chart
            ch.Series(0).Points.AddXY(s_points(h, 0), s_points(h, 1))   'stage #1 Loss [%]
            ch.Series(1).Points.AddXY(s_points(h, 0), s_points(h, 2))   'stage #2 Loss [%]
        Next h

        ch.Series(0).Points(6).Label = "Cyclone #1"
        ch.Series(1).Points(4).Label = "Cyclone #2"
    End Sub
    Private Sub Draw_chart3(ch As Chart)
        'Lust Load correction formula
        Dim s_points(150, 2) As Double
        Dim h As Integer

        ch.Series.Clear()
        ch.ChartAreas.Clear()
        ch.Titles.Clear()
        ch.ChartAreas.Add("ChartArea0")

        ch.Series.Add("Series1")
        ch.Series(0).ChartArea = "ChartArea0"
        ch.Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
        ch.Series(0).BorderWidth = 2
        ch.Series(0).IsVisibleInLegend = False

        ch.Titles.Add("Dust Load correction factor")
        ch.ChartAreas("ChartArea0").AxisX.Title = "Dust load [gr/Am3]"
        ch.ChartAreas("ChartArea0").AxisY.Title = "Dust Load correction factor"
        ch.ChartAreas("ChartArea0").AxisY.Minimum = 1               'Correction factor
        ch.ChartAreas("ChartArea0").AxisY.Maximum = 2               'Correction factor
        ch.ChartAreas("ChartArea0").AxisX.Minimum = 0               'Dustload
        ch.ChartAreas("ChartArea0").AxisX.Maximum = 150             'Dustload

        '----- now calc chart points #1 and #2 --------------------------
        s_points(0, 0) = 0                                          'Dust load 0 [kg/Am3]
        s_points(0, 1) = 1                                          'Correction factor [-]

        For h = 1 To 150                                            'Dust load [kg/Am3]
            s_points(h, 0) = h                                      'Dust load [kg/Am3]
            s_points(h, 1) = Calc_dust_load_correction(h / 1000)    'Correction factor [-]
        Next

        '------ now present-------------
        For h = 0 To 150 - 1   'Fill line chart
            ch.Series(0).Points.AddXY(s_points(h, 0), s_points(h, 1))   '
        Next h

        ch.Series(0).Points(6).Label = "Load correction"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles TabPage9.Enter, CheckBox6.CheckedChanged, CheckBox7.CheckedChanged, CheckBox4.CheckedChanged, CheckBox9.CheckedChanged, CheckBox8.CheckedChanged, CheckBox10.CheckedChanged, CheckBox5.CheckedChanged, CheckBox12.CheckedChanged, CheckBox11.CheckedChanged, CheckBox1.CheckedChanged, CheckBox13.CheckedChanged
        Calc_sequence()
    End Sub
    Private Sub Calc_sequence()
        Dim case_nr As Integer = CInt(NumericUpDown30.Value)

        If ComboBox1.SelectedIndex > -1 And ComboBox2.SelectedIndex > -1 Then
            Dust_load_correction(case_nr)
            Get_input_calc_1(case_nr)       'This is the CASE number
            Calc_part_dia_loss(case_nr)

            Calc_stage1(case_nr)            'Calc according stage #1
            Calc_stage2(case_nr)            'Calc according stage #2

            Calc_stage1(case_nr)            'Calc according stage #1
            Calc_stage2(case_nr)            'Calc according stage #2

            Calc_stage1(case_nr)            'Calc according stage #1
            Calc_stage2(case_nr)            'Calc according stage #2
            Calc_stage1_2_comb()            'Calc stage #1 and stage #2 combined

            Present_loss_grid1()            'Present the results stage #1
            Present_loss_grid2()            'Present the results stage #2
            Present_Datagridview1(case_nr)  'Present the results stage #1

            Draw_chart1(Chart1)             'Present the results 
            Draw_chart2(Chart2)             'Present the results loss curve
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Save the project data to file
        'Return to case 0, save data to file, return to previous selected case

        If TextBox28.Text.Trim.Length > 0 And TextBox29.Text.Trim.Length > 0 Then
            Save_present_case_to_array()            'Store data in array
            Save_to_disk()                          'Store project on disk
        Else
            MessageBox.Show("Complete Quote and Tag number")
        End If
    End Sub
    Private Sub Save_to_disk()
        Dim filename, user As String
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        'TextBox24.Text &= "line 1162, save project to disk" & vbCrLf
        If TextBox28.Text.Trim.Length = 0 Or TextBox29.Text.Trim.Length = 0 Then
            MessageBox.Show("Complete Quote and Tag number")
            Exit Sub
        End If

        Save_present_case_to_array()            'Store data in array

        '------------- create filemame ----------
        user = Trim(Environment.UserName)           'User name on the screen
        filename = "Cyclone_select_" & TextBox28.Text & "_" & TextBox29.Text & DateTime.Now.ToString("_yyyy_MM_dd_") & user & ".vtk2"
        filename = Replace(filename, Chr(32), Chr(95)) 'Replace the space's

        '------------- Create directories if they do not exist ----
        Try
            If (Not System.IO.Directory.Exists(dirpath_Temp)) Then System.IO.Directory.CreateDirectory(dirpath_Temp)
            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
        Catch ex As Exception
            MessageBox.Show("Can not create directory on the VTK intranet (L6286) " & vbCrLf & vbCrLf & ex.Message)
        End Try

        Try
            If Directory.Exists(dirpath_Eng) Then
                filename = dirpath_Eng & filename 'used at VTK with intranet
            Else
                filename = dirpath_tmp & filename 'used at VTK with intranet'used at home
                MessageBox.Show("VTK intranet not acceasable, save on " & filename)
            End If

            '--- Delete existing file -------
            If System.IO.File.Exists(filename) = True Then
                System.IO.File.Delete(filename)
            End If

            '--- and save new file to disk -------
            Dim fStream As New FileStream(filename, FileMode.CreateNew)
            bf.Serialize(fStream, _cees) ' write to file
            fStream.Close()
        Catch ex As Exception
            MessageBox.Show("Line 6298, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    Private Sub Retrieve_from_disk()
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        OpenFileDialog1.FileName = "Cyclone_select_*"

        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_tmp  'used at home
        End If

        OpenFileDialog1.Title = "Open a Text File"
        OpenFileDialog1.Filter = "VTK2 Files|*.vtk2|VTK1 file|*.vtk"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                Dim fStream As New FileStream(OpenFileDialog1.FileName, FileMode.Open) With {
                    .Position = 0 ' reset stream pointer
                    }
                _cees = CType(bf.Deserialize(fStream), Input_struct()) ' read from file
                fStream.Close()
            Catch ex As Exception
                MessageBox.Show("Line 1013, " & ex.Message)  ' Show the exception's message.
            End Try
            TextBox24.Text &= "line 1220, Retrieved  project from disk" & vbCrLf
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Retrieve project from disk and goto case nr 0

        Retrieve_from_disk()        'Read from disk to array
        Fill_Screen_from_array(0)   'Refresh screen data case 0
        Calc_sequence()
    End Sub

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
            Write_to_word_com()     'Commercial data to Word
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

        Dim chart_size As Integer = 65  '% of original picture size
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
            oPara1.Range.Text = "VTK Sales, cyclone single stage"
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
            oTable.Rows(1).Range.Font.Bold = CInt(True)
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
            oTable.Cell(row, 2).Range.Text = NumericUpDown1.Value.ToString("F0")
            oTable.Cell(row, 3).Range.Text = "[Am3/hr]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Temperature"
            oTable.Cell(row, 2).Range.Text = NumericUpDown18.Value.ToString("F0")
            oTable.Cell(row, 3).Range.Text = "[c]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Inlet pressure"
            oTable.Cell(row, 2).Range.Text = NumericUpDown19.Value.ToString("F1")
            oTable.Cell(row, 3).Range.Text = "[mbar abs]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Particle density "
            oTable.Cell(row, 2).Range.Text = numericUpDown2.Value.ToString("F0")
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Gas density "
            oTable.Cell(row, 2).Range.Text = numericUpDown3.Value.ToString("F3")
            oTable.Cell(row, 3).Range.Text = "[kg/m3]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Air viscosity"
            oTable.Cell(row, 2).Range.Text = numericUpDown14.Value.ToString("F3")
            oTable.Cell(row, 3).Range.Text = "[centi Poise]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Dust load"
            oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value.ToString("F2")
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
            oTable.Cell(row, 3).Range.Text = "[%]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(1)
            oTable.Columns(3).Width = oWord.InchesToPoints(2)

            oTable.Rows(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------cyclone data-------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Cyclone data"
            row += 1
            '----------- stage #1 ---------------
            oTable.Cell(row, 1).Range.Text = "Cyclone type stage #1 "
            oTable.Cell(row, 2).Range.Text = ComboBox1.SelectedItem.ToString
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Body diameter #1"
            oTable.Cell(row, 2).Range.Text = numericUpDown5.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "No parallel #1"
            oTable.Cell(row, 2).Range.Text = NumericUpDown20.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            '----------- stage #2 ---------------
            oTable.Cell(row, 1).Range.Text = "Cyclone type stage #2 "
            oTable.Cell(row, 2).Range.Text = ComboBox2.SelectedItem.ToString
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Body diameter #2"
            oTable.Cell(row, 2).Range.Text = NumericUpDown34.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "No parallel #2"
            oTable.Cell(row, 2).Range.Text = NumericUpDown33.Value.ToString
            oTable.Cell(row, 3).Range.Text = "[-]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(1)
            oTable.Columns(3).Width = oWord.InchesToPoints(2)

            oTable.Rows(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------Process data-------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Process data"
            row += 1
            '----------- stage #1 ---------------
            oTable.Cell(row, 1).Range.Text = "Inlet speed #1 "
            oTable.Cell(row, 2).Range.Text = TextBox16.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Outlet speed #1"
            oTable.Cell(row, 2).Range.Text = TextBox22.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Pressure loss #1"
            oTable.Cell(row, 2).Range.Text = TextBox17.Text
            oTable.Cell(row, 3).Range.Text = "[Pa]"
            row += 1
            '----------- stage #2 ---------------
            oTable.Cell(row, 1).Range.Text = "Inlet speed #2 "
            oTable.Cell(row, 2).Range.Text = TextBox80.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Outlet speed #2"
            oTable.Cell(row, 2).Range.Text = TextBox77.Text
            oTable.Cell(row, 3).Range.Text = "[m/s]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Pressure loss #2"
            oTable.Cell(row, 2).Range.Text = TextBox79.Text
            oTable.Cell(row, 3).Range.Text = "[Pa]"


            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(1)
            oTable.Columns(3).Width = oWord.InchesToPoints(2)

            oTable.Rows(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '---------------Calculation date stage #1-------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 24, 10)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Stage #1"
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
                oTable.Cell(row, 1).Range.Text = CType(DataGridView1.Rows(j).Cells(0).Value, String)
                oTable.Cell(row, 2).Range.Text = CType(DataGridView1.Rows(j).Cells(1).Value, String)
                oTable.Cell(row, 3).Range.Text = CType(DataGridView1.Rows(j).Cells(2).Value, String)
                oTable.Cell(row, 4).Range.Text = CType(DataGridView1.Rows(j).Cells(3).Value, String)
                oTable.Cell(row, 5).Range.Text = CType(DataGridView1.Rows(j).Cells(4).Value, String)
                oTable.Cell(row, 6).Range.Text = CType(DataGridView1.Rows(j).Cells(5).Value, String)
                oTable.Cell(row, 7).Range.Text = CType(DataGridView1.Rows(j).Cells(6).Value, String)
                oTable.Cell(row, 8).Range.Text = CType(DataGridView1.Rows(j).Cells(7).Value, String)
                oTable.Cell(row, 9).Range.Text = CType(DataGridView1.Rows(j).Cells(8).Value, String)
            Next

            For j = 1 To 8
                oTable.Columns(j).Width = oWord.InchesToPoints(0.75)   'Change width of columns 
            Next
            oTable.Rows(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            ''---------------Calculation date stage #2-------------------------------
            ''Insert a table, fill it with data and change the column widths.
            'oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 24, 10)
            'oTable.Range.ParagraphFormat.SpaceAfter = 1
            'oTable.Range.Font.Size = 10
            'oTable.Range.Font.Bold = CInt(False)
            'oTable.Rows(1).Range.Font.Bold = CInt(True)
            'row = 1
            'oTable.Cell(row, 1).Range.Text = "Stage #2"
            'row += 1

            'oTable.Cell(row, 1).Range.Text = "Dia class[mu]"
            'oTable.Cell(row, 2).Range.Text = "Feed psd cumm [%]"
            'oTable.Cell(row, 3).Range.Text = "Feed psd diff [%]"
            'oTable.Cell(row, 4).Range.Text = "Loss of feed [%]"
            'oTable.Cell(row, 5).Range.Text = "Loss abs [%]"
            'oTable.Cell(row, 6).Range.Text = "Loss cum [%]"
            'oTable.Cell(row, 7).Range.Text = "Catch abs [%]"
            'oTable.Cell(row, 8).Range.Text = "Catch cum [%]"
            'oTable.Cell(row, 9).Range.Text = "Efficiency [%]"

            'For j = 0 To 22
            '    row += 1
            '    oTable.Cell(row, 1).Range.Text = CType(DataGridView1.Rows(j).Cells(0).Value, String)
            '    oTable.Cell(row, 2).Range.Text = CType(DataGridView1.Rows(j).Cells(1).Value, String)
            '    oTable.Cell(row, 3).Range.Text = CType(DataGridView1.Rows(j).Cells(2).Value, String)
            '    oTable.Cell(row, 4).Range.Text = CType(DataGridView1.Rows(j).Cells(3).Value, String)
            '    oTable.Cell(row, 5).Range.Text = CType(DataGridView1.Rows(j).Cells(4).Value, String)
            '    oTable.Cell(row, 6).Range.Text = CType(DataGridView1.Rows(j).Cells(5).Value, String)
            '    oTable.Cell(row, 7).Range.Text = CType(DataGridView1.Rows(j).Cells(6).Value, String)
            '    oTable.Cell(row, 8).Range.Text = CType(DataGridView1.Rows(j).Cells(7).Value, String)
            '    oTable.Cell(row, 9).Range.Text = CType(DataGridView1.Rows(j).Cells(8).Value, String)
            'Next

            'For j = 1 To 8
            '    oTable.Columns(j).Width = oWord.InchesToPoints(0.75)   'Change width of columns 
            'Next
            'oTable.Rows(1).Range.Font.Bold = CInt(True)
            'oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '------------------save Chart2 (Loss curve)---------------- 
            Draw_chart2(Chart2)
            file_name = dirpath_Temp & "Chart_loss.Jpeg"
            'Chart2.SaveImage(file_name, System.Drawing.Imaging.ImageFormat.Jpeg)
            Chart1.SaveImage(file_name, System.Drawing.Imaging.ImageFormat.Jpeg)
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
        TextBox30.Text = Visco.ToString("F5")   'centi Poise
        TextBox146.Text = (Visco * 0.001).ToString("F7")  'Pa.sec
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

        DataGridView2.EnableHeadersVisualStyles = False                     'For backcolor
        DataGridView2.Columns(0).HeaderText = "Dia upper [mu]"              '
        DataGridView2.Columns(1).HeaderText = "Dia (aver.) class [mu]"      '         
        DataGridView2.Columns(2).HeaderText = "Dia/k [-]"                   '
        DataGridView2.Columns(3).HeaderText = "Loss overall [-]"            '
        DataGridView2.Columns(4).HeaderText = "Loss overall Corrected"      '
        DataGridView2.Columns(5).HeaderText = "Grade eff. (st1) [%]"        'Catch [%]
        DataGridView2.Columns(6).HeaderText = "Group no [-]"                '
        DataGridView2.Columns(7).HeaderText = "d1 lower dia [mu]"           '
        DataGridView2.Columns(8).HeaderText = "d2 upper dia [mu]"           '
        DataGridView2.Columns(9).HeaderText = "p1 input [%]"                '
        DataGridView2.Columns(10).HeaderText = "p2 input [%]"               '
        DataGridView2.Columns(11).HeaderText = "k [-]"                      '    
        DataGridView2.Columns(12).HeaderText = "m [-]"                      '
        DataGridView2.Columns(13).HeaderText = "interpol. psd cum [-]"      '
        DataGridView2.Columns(14).HeaderText = "psd cum [%]"                '
        DataGridView2.Columns(15).HeaderText = "psd diff [%]"               '
        DataGridView2.Columns(16).HeaderText = "loss abs [%]"               '
        DataGridView2.Columns(17).HeaderText = "loss corr abs [%]"          '

        DataGridView2.Columns(1).HeaderCell.Style.BackColor = Color.Yellow  'Chart
        DataGridView2.Columns(5).HeaderCell.Style.BackColor = Color.Yellow  'Chart

        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        'DataGridView2.AutoSize = False
        For row = 1 To 110  'Fill the DataGrid
            j = row - 1
            DataGridView2.Rows(j).Cells(0).Value = _cees(ks).stage1(j).dia.ToString("F8")
            DataGridView2.Rows(j).Cells(1).Value = _cees(ks).stage1(j).d_ave.ToString("F8")          'Average diameter
            DataGridView2.Rows(j).Cells(2).Value = _cees(ks).stage1(j).d_ave_K.ToString("F5")        'Average dia/K stokes
            DataGridView2.Rows(j).Cells(3).Value = _cees(ks).stage1(j).loss_overall.ToString("F5")   'Loss 
            DataGridView2.Rows(j).Cells(4).Value = _cees(ks).stage1(j).loss_overall_C.ToString("F5") 'Loss 
            DataGridView2.Rows(j).Cells(5).Value = _cees(ks).stage1(j).catch_chart.ToString("F5")    'Catch
            DataGridView2.Rows(j).Cells(6).Value = _cees(ks).stage1(j).i_grp.ToString              'Groep nummer
            DataGridView2.Rows(j).Cells(7).Value = _cees(ks).stage1(j).i_d1.ToString("F5")         'class lower dia limit
            DataGridView2.Rows(j).Cells(8).Value = _cees(ks).stage1(j).i_d2.ToString("F5")         'class upper dia limit
            DataGridView2.Rows(j).Cells(9).Value = _cees(ks).stage1(j).i_p1.ToString("F5")         'User input percentage
            DataGridView2.Rows(j).Cells(10).Value = _cees(ks).stage1(j).i_p2.ToString("F5")        'User input percentage
            DataGridView2.Rows(j).Cells(11).Value = _cees(ks).stage1(j).i_k.ToString("F3")         'k [-]
            DataGridView2.Rows(j).Cells(12).Value = _cees(ks).stage1(j).i_m.ToString("F5")         'm [-]
            DataGridView2.Rows(j).Cells(13).Value = _cees(ks).stage1(j).psd_cum.ToString("F5")     'interpol. psd cum [-]
            DataGridView2.Rows(j).Cells(14).Value = _cees(ks).stage1(j).psd_cum_pro.ToString("F5") '[%] psd cum
            DataGridView2.Rows(j).Cells(15).Value = _cees(ks).stage1(j).psd_dif.ToString("F5")     '[%] psd diff
            DataGridView2.Rows(j).Cells(16).Value = _cees(ks).stage1(j).loss_abs.ToString("F5")    '[%] loss abs
            DataGridView2.Rows(j).Cells(17).Value = _cees(ks).stage1(j).loss_abs_C.ToString("F5")  '[%] loss corr abs 
        Next
        DataGridView2.Rows(111).Cells(15).Value = _cees(ks).sum_psd_diff1.ToString("F5")  'total_psd_diff.
        DataGridView2.Rows(111).Cells(16).Value = _cees(ks).sum_loss1.ToString("F5") 'total_abs_loss.ToString("F5")
        DataGridView2.Rows(111).Cells(17).Value = _cees(ks).sum_loss_C1.ToString("F5") 'total_abs_loss_C.ToString("F5")
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

        DataGridView3.EnableHeadersVisualStyles = False                         'For backcolor
        DataGridView3.Columns(0).HeaderText = "Dia upper [mu]"                  '
        DataGridView3.Columns(1).HeaderText = "Dia (aver.) class [mu]"          '
        DataGridView3.Columns(2).HeaderText = "Dia/k [-]"                       '
        DataGridView3.Columns(3).HeaderText = "Loss overall [-]"                '
        DataGridView3.Columns(4).HeaderText = "Loss overall Corrected [-]"      '
        DataGridView3.Columns(5).HeaderText = "Grade eff. (st2) [%]"            'Catch chart [%]
        DataGridView3.Columns(6).HeaderText = "Group no [-]"                    '
        DataGridView3.Columns(7).HeaderText = "d1 lower dia [mu]"               '
        DataGridView3.Columns(8).HeaderText = "d2 upper dia [mu]"               '
        DataGridView3.Columns(9).HeaderText = "p1 input [%]"                    '
        DataGridView3.Columns(10).HeaderText = "p2 input [%]"                   '
        DataGridView3.Columns(11).HeaderText = "k [-]"                          '    
        DataGridView3.Columns(12).HeaderText = "m [-]"                          '
        DataGridView3.Columns(13).HeaderText = "interpol. psd cum [-]"          '
        DataGridView3.Columns(14).HeaderText = "psd_cum_pro for chart"          '
        DataGridView3.Columns(15).HeaderText = "psd diff [%] of 1-stage"        '
        DataGridView3.Columns(16).HeaderText = "psd diff [%] of 2-stage"        '
        DataGridView3.Columns(17).HeaderText = "catch loss abs [%]"             '
        DataGridView3.Columns(18).HeaderText = "loss abs corrected [%]"         '

        DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        DataGridView3.Columns(1).HeaderCell.Style.BackColor = Color.Yellow      'For chart
        DataGridView3.Columns(5).HeaderCell.Style.BackColor = Color.Yellow      'For chart
        DataGridView3.Columns(14).HeaderCell.Style.BackColor = Color.Yellow     'For chart

        For row = 1 To 110  'Fill the DataGrid
            j = row - 1
            DataGridView3.Rows(j).Cells(0).Value = _cees(ks).stage2(j).dia.ToString("F8")          'Dia particle
            DataGridView3.Rows(j).Cells(1).Value = _cees(ks).stage2(j).d_ave.ToString("F8")        'Average diameter
            DataGridView3.Rows(j).Cells(2).Value = _cees(ks).stage2(j).d_ave_K.ToString("F5")      'Average dia/K stokes
            DataGridView3.Rows(j).Cells(3).Value = _cees(ks).stage2(j).loss_overall.ToString("F5") 'Loss 
            DataGridView3.Rows(j).Cells(4).Value = _cees(ks).stage2(j).loss_overall_C.ToString("F5") 'Loss corrected 
            DataGridView3.Rows(j).Cells(5).Value = _cees(ks).stage2(j).catch_chart.ToString("F5")  'Catch for chart
            DataGridView3.Rows(j).Cells(6).Value = _cees(ks).stage2(j).i_grp.ToString              'Groep nummer
            DataGridView3.Rows(j).Cells(7).Value = _cees(ks).stage2(j).i_d1.ToString("F5")         'class lower dia limit
            DataGridView3.Rows(j).Cells(8).Value = _cees(ks).stage2(j).i_d2.ToString("F5")         'class upper dia limit
            DataGridView3.Rows(j).Cells(9).Value = _cees(ks).stage2(j).i_p1.ToString("F5")         'User lower input percentage
            DataGridView3.Rows(j).Cells(10).Value = _cees(ks).stage2(j).i_p2.ToString("F5")        'User upper input percentage
            DataGridView3.Rows(j).Cells(11).Value = _cees(ks).stage2(j).i_k.ToString("F5")         'parameter k
            DataGridView3.Rows(j).Cells(12).Value = _cees(ks).stage2(j).i_m.ToString("F5")         'parameter m
            DataGridView3.Rows(j).Cells(13).Value = _cees(ks).stage2(j).psd_cum.ToString("F5")     '[-] interpol. psd cum
            DataGridView3.Rows(j).Cells(14).Value = _cees(ks).stage2(j).psd_cum_pro.ToString("F5") '[%] interpol. psd cum x100
            DataGridView3.Rows(j).Cells(15).Value = _cees(ks).stage1(j).psd_dif.ToString("F5")     '[%] psd diff of 1-stage
            DataGridView3.Rows(j).Cells(16).Value = _cees(ks).stage2(j).psd_dif.ToString("F5")     '[%] psd diff of 2-stage
            DataGridView3.Rows(j).Cells(17).Value = _cees(ks).stage2(j).loss_abs.ToString("F5")    '[%] loss abs 
            DataGridView3.Rows(j).Cells(18).Value = _cees(ks).stage2(j).loss_abs_C.ToString("F5")  '[%] loss abs corrected
        Next
        DataGridView3.Rows(111).Cells(15).Value = _cees(ks).sum_psd_diff1.ToString("F5")   'total_psd_diff.
        DataGridView3.Rows(111).Cells(16).Value = _cees(ks).sum_psd_diff2.ToString("F5")   'total_psd_diff.
        DataGridView3.Rows(111).Cells(17).Value = _cees(ks).sum_loss2.ToString("F5")       'total_abs_loss.ToString("F5")
        DataGridView3.Rows(111).Cells(18).Value = _cees(ks).sum_loss_C2.ToString("F5")     'total_abs_loss_C.ToString("F5")
    End Sub

    Private Sub Calc_k_and_m(ByRef g As GvG_Calc_struct)
        Dim k, m As Double
        'k and m are based on particle diameter and percentages

        k = Log(Log(g.i_p1) / Log(g.i_p2)) / Log(g.i_d1 / g.i_d2)   '====== k ===========
        m = g.i_d1 / ((-Log(g.i_p1)) ^ (1 / g.i_k))                 '====== m ===========

        '==== preventing errors =====
        If Double.IsNaN(k) Then k = 1
        If Double.IsNaN(m) Then m = 1

        g.i_k = k
        g.i_m = m
    End Sub

    Private Sub Calc_stage1(ks As Integer)
        'This is the standard VTK cyclone calculation for case "ks" 
        Dim i As Integer = 0
        Dim perc_smallest_part1 As Double
        Dim fac_m As Double
        Dim words() As String
        Dim density_ratio1, density_ratio2 As Double

        If Double.IsNaN(_cees(ks).stage1(0).dia) Or Double.IsInfinity(_cees(ks).stage1(0).dia) Then Exit Sub

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

        _cees(ks).stage1(0).psd_cum = Math.E ^ (-((_cees(ks).stage1(i).dia / _cees(ks).stage1(i).i_m) ^ _cees(ks).stage1(i).i_k))
        _cees(ks).stage1(0).psd_cum_pro = _cees(ks).stage1(i).psd_cum * 100

        _cees(ks).stage1(0).psd_dif = 100 * (1 - _cees(ks).stage1(i).psd_cum)
        _cees(ks).stage1(0).loss_abs = _cees(ks).stage1(i).loss_overall * _cees(ks).stage1(i).psd_dif
        _cees(ks).stage1(0).loss_abs_C = _cees(ks).stage1(i).loss_overall_C * _cees(ks).stage1(i).psd_dif

        '----- initial values --------
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

        '------------ Particle diameter calculation step -----
        _istep = (_cees(ks).Dmax1 / _cees(ks).Dmin2) ^ (1 / 110) 'Calculation step

        For i = 1 To 110
            _cees(ks).stage1(i).dia = _cees(ks).stage1(i - 1).dia * _istep
            _cees(ks).stage1(i).d_ave = (_cees(ks).stage1(i - 1).dia + _cees(ks).stage1(i).dia) / 2       'Average diameter
            _cees(ks).stage1(i).d_ave_K = _cees(ks).stage1(i).d_ave / _cees(ks).Kstokes1        'dia/k_stokes
            _cees(ks).stage1(i).loss_overall = Calc_verlies(_cees(ks).stage1(i).d_ave, False, _cees(ks).Kstokes1, 1)   '[-] loss overall
            Calc_verlies_corrected(_cees(ks).stage1(i), 1)                               '[-] loss overall corrected


            'If _cees(ks).stage1(i).loss_overall_C <> _cees(ks).stage1(i).loss_overall Then
            '    MsgBox(_cees(ks).stage1(i).loss_overall_C.ToString & ",  " & _cees(ks).stage1(i).loss_overall.ToString)
            'End If

            If CheckBox2.Checked Then
                _cees(ks).stage1(i).catch_chart = (1 - _cees(ks).stage1(i).loss_overall_C) * 100 '[%] Corrected
            Else
                _cees(ks).stage1(i).catch_chart = (1 - _cees(ks).stage1(i).loss_overall) * 100    '[%] NOT corrected
            End If
            Size_classification(_cees(ks).stage1(i))                                   'Classify this part size

            '====to prevent silly results====
            If _cees(ks).stage1(i).i_grp <> 11 Then
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

            '----- sum value incremental values -----
            _cees(ks).sum_psd_diff1 += _cees(ks).stage1(i).psd_dif
            _cees(ks).sum_loss1 += _cees(ks).stage1(i).loss_abs
            _cees(ks).sum_loss_C1 += _cees(ks).stage1(i).loss_abs_C
        Next
        _cees(ks).loss_total1 = _cees(ks).sum_loss_C1 + ((100 - _cees(ks).sum_psd_diff1) * perc_smallest_part1)

        _cees(ks).emmis1_Am3 = NumericUpDown4.Value * (_cees(ks).loss_total1 / 100)  '[g/Am3]
        _cees(ks).emmis1_Nm3 = _cees(ks).emmis1_Am3 * Calc_Normal_density(_cees(ks).Ro_gas1_Am3, _cees(ks).p1_abs, _cees(ks).Temp)

        '----------Dust load stage #2 in emission stage #1 -----------
        _cees(ks).dust2_Am3 = _cees(ks).emmis1_Am3

        '--------- Density ratio stage #2 kg/Nm3 and kg/Am3 ---------
        density_ratio2 = _cees(ks).Ro_gas2_Nm3 / _cees(ks).Ro_gas2_Am3  '[-]
        _cees(ks).dust2_Nm3 = _cees(ks).dust2_Am3 * density_ratio2      'Dust load [gram/Nm3]

        CheckBox3.Checked = CBool(IIf(_cees(ks).dust2_Am3 > 20, True, False))
        _cees(ks).Efficiency1 = 100 - _cees(ks).loss_total1        '[%] Efficiency

        '----------- present stage #1-----------
        TextBox51.Text = _cees(ks).Dmax1.ToString("F2")     'diameter [mu] 100% catch
        TextBox52.Text = _cees(ks).Dmin1.ToString("F2")     'diameter [mu] 100% loss
        TextBox56.Text = ComboBox1.Text                     'Cyclone type
        TextBox57.Text = CheckBox2.Checked.ToString         'Correction stage #1
        TextBox70.Text = _cees(ks).dust2_Am3.ToString("F2") 'Dust load [gram/Am3]

        TextBox118.Text = _cees(ks).sum_psd_diff1.ToString("F3")
        TextBox54.Text = _cees(ks).sum_loss1.ToString("F3")
        TextBox34.Text = _cees(ks).sum_loss_C1.ToString("F3")

        '---------Density ratio stage#1  kg/Nm3 and kg/Am3 ---------
        density_ratio1 = _cees(ks).Ro_gas1_Nm3 / _cees(ks).Ro_gas1_Am3  '[-]

        '----------- Dust load correction stage #1 ------------------
        If CheckBox2.Checked Then
            _cees(ks).emmis1_Am3 = _cees(ks).emmis1_Am3
            _cees(ks).emmis1_Nm3 = _cees(ks).emmis1_Am3 * density_ratio1
            TextBox58.Text = _cees(ks).loss_total1.ToString("F3")    '[%] Corrected 
        Else
            _cees(ks).emmis1_Am3 = (NumericUpDown4.Value * _cees(ks).sum_loss1 / 100)
            _cees(ks).emmis1_Nm3 = _cees(ks).emmis1_Am3 * density_ratio1
            TextBox58.Text = _cees(ks).sum_loss1.ToString("F3")      '[%] NOT Corrected  
        End If

        TextBox59.Text = _cees(ks).Efficiency1.ToString("F3")
        TextBox21.Text = _cees(ks).Efficiency1.ToString("F3")
        TextBox60.Text = _cees(ks).emmis1_Am3.ToString("F3")
        TextBox18.Text = _cees(ks).emmis1_Am3.ToString("F3")
        TextBox133.Text = _cees(ks).emmis1_Nm3.ToString("F3")
    End Sub

    Private Sub Calc_stage2(ks As Integer)
        'This is the standard VTK cyclone calculation 
        Dim i As Integer = 0
        Dim perc_smallest_part2 As Double
        Dim fac_m As Double
        Dim words() As String
        Dim kgh, tot_kgh As Double
        Dim Eff_comb As Double      'Efficiency stage #1 and #2

        If Double.IsNaN(_cees(ks).stage2(0).dia) Or Double.IsInfinity(_cees(ks).stage2(0).dia) Then Exit Sub

        '----------- stof belasting ------------
        tot_kgh = _cees(ks).Flow2 * _cees(ks).dust2_Am3 / 1000 * 3600 * _cees(ks).Noc2     '[kg/hr] Dust inlet 
        kgh = tot_kgh / _cees(ks).Noc2                          '[kg/hr/Cy] Dust inlet 

        TextBox100.Text = kgh.ToString("F0")
        TextBox101.Text = tot_kgh.ToString("F0")

        '--------- now the particles (====Grid line 0======)------------
        _cees(ks).stage2(0).dia = _cees(ks).stage1(0).dia                                   'Copy stage #1
        _cees(ks).stage2(0).d_ave = _cees(ks).stage2(0).dia / 2                             'Average diameter
        _cees(ks).stage2(0).d_ave_K = _cees(ks).stage2(0).d_ave / _cees(ks).Kstokes2        'dia/k_stokes
        _cees(ks).stage2(0).loss_overall = Calc_verlies(_cees(ks).stage2(0).d_ave_K, False, _cees(ks).Kstokes2, 2)     '[-] loss overall
        Calc_verlies_corrected(_cees(ks).stage2(0), 2)                               '[-] loss overall corrected
        _cees(ks).stage2(0).catch_chart = (1 - _cees(ks).stage2(0).loss_overall_C) * 100    '[%]
        Size_classification(_cees(ks).stage2(0))                                  'groepnummer

        If _cees(ks).stage2(i).i_grp <> 11 Then
            Calc_k_and_m(_cees(ks).stage2(0))
            _cees(ks).stage2(0).psd_cum = Math.E ^ (-((_cees(ks).stage2(0).dia / _cees(ks).stage2(0).i_m) ^ _cees(ks).stage2(0).i_k))
            _cees(ks).stage2(0).psd_cum_pro = _cees(ks).stage2(0).psd_cum * 100 '[%]
            _cees(ks).stage2(0).psd_dif = 100 * _cees(ks).stage1(0).loss_abs / (100 - _cees(ks).Efficiency1)              'LOSS STAGE #1
        Else
            _cees(ks).stage2(0).i_k = 0
            _cees(ks).stage2(0).i_m = 0
            _cees(ks).stage2(0).psd_cum = 0
            _cees(ks).stage2(0).psd_cum_pro = 0
            _cees(ks).stage2(0).psd_dif = 0
        End If

        _cees(ks).stage2(0).loss_abs = _cees(ks).stage2(0).loss_overall * _cees(ks).stage2(0).psd_dif
        _cees(ks).stage2(0).loss_abs_C = _cees(ks).stage2(0).loss_overall_C * _cees(ks).stage2(0).psd_dif

        '----- initial values -------
        _cees(ks).sum_psd_diff2 = 0 '_cees(ks).stage2(0).psd_dif
        _cees(ks).sum_loss2 = 0 ' _cees(ks).stage2(0).loss_abs
        _cees(ks).sum_loss_C2 = 0 ' _cees(ks).stage2(0).loss_abs_C

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
        perc_smallest_part2 = 0.0000001                    'smallest particle [%]
        _cees(ks).Dmax2 = Calc_dia_particle(perc_smallest_part2, _cees(ks).Kstokes2, 2)     '=100% loss (biggest particle)
        _cees(ks).Dmin2 = _cees(ks).Kstokes2 * fac_m                'diameter smallest particle caught


        For i = 1 To 110    '=========Stage #2, Grid lines 1...============ 
            If Double.IsNaN(_cees(ks).stage1(ks).dia) Or Double.IsInfinity(_cees(ks).stage1(ks).dia) Then Exit Sub

            _cees(ks).stage2(i).dia = _cees(ks).stage1(i).dia        'Diameter Copy stage #1
            _cees(ks).stage2(i).d_ave = _cees(ks).stage1(i).d_ave            'Average diameter
            _cees(ks).stage2(i).d_ave_K = _cees(ks).stage2(i).d_ave / _cees(ks).Kstokes2          'dia/k_stokes
            _cees(ks).stage2(i).loss_overall = Calc_verlies(_cees(ks).stage2(i).d_ave, False, _cees(ks).Kstokes2, 2)   '[-] loss overall
            Calc_verlies_corrected(_cees(ks).stage2(i), 2)                                '[-] loss overall corrected

            '------------- Load correction stage #2 -------------------
            If CheckBox3.Checked Then
                _cees(ks).stage2(i).catch_chart = (1 - _cees(ks).stage2(i).loss_overall_C) * 100  '[%] Corrected
            Else
                _cees(ks).stage2(i).catch_chart = (1 - _cees(ks).stage2(i).loss_overall) * 100    '[%] NOT corrected
            End If


            Size_classification(_cees(ks).stage2(i))                                     'Calc
            _cees(ks).stage2(i).i_grp = _cees(ks).stage1(i).i_grp

            If _cees(ks).stage2(i).i_grp <> 11 Then
                Calc_k_and_m(_cees(ks).stage2(i))
                _cees(ks).stage2(i).psd_cum = Math.E ^ (-(_cees(ks).stage2(i).dia / _cees(ks).stage2(i).i_m) ^ _cees(ks).stage2(i).i_k)
                _cees(ks).stage2(i).psd_cum_pro = _cees(ks).stage2(i).psd_cum * 100 '[%]
                _cees(ks).stage2(i).psd_dif = 100 * _cees(ks).stage1(i).loss_abs_C / _cees(ks).sum_loss_C1
            Else
                _cees(ks).stage2(i).i_k = 0
                _cees(ks).stage2(i).i_m = 0
                _cees(ks).stage2(i).psd_cum = 0
                _cees(ks).stage2(i).psd_cum_pro = 0
                _cees(ks).stage2(i).psd_dif = 0
            End If

            _cees(ks).stage2(i).loss_abs = _cees(ks).stage2(i).loss_overall * _cees(ks).stage2(i).psd_dif
            _cees(ks).stage2(i).loss_abs_C = _cees(ks).stage2(i).loss_overall_C * _cees(ks).stage2(i).psd_dif

            '----- sum value incremental values -----
            _cees(ks).sum_psd_diff2 += _cees(ks).stage2(i).psd_dif
            _cees(ks).sum_loss2 += _cees(ks).stage2(i).loss_abs
            _cees(ks).sum_loss_C2 += _cees(ks).stage2(i).loss_abs_C
        Next i
        _cees(ks).loss_total2 = _cees(ks).sum_loss_C2 + ((100 - _cees(ks).sum_psd_diff2) * perc_smallest_part2)
        _cees(ks).emmis2_Am3 = _cees(ks).emmis1_Am3 * _cees(ks).loss_total2 / 100
        _cees(ks).Efficiency2 = 100 - _cees(ks).loss_total2      '[%] Efficiency

        '------ combined efficiency -----
        Eff_comb = _cees(ks).Efficiency1 + (1 - _cees(ks).Efficiency1 / 100) * _cees(ks).Efficiency2


        '----------- present stage #2 -----------
        TextBox63.Text = ComboBox2.Text                     'Cyclone type
        TextBox64.Text = CheckBox3.Checked.ToString         'Hi load correction stage #2
        TextBox110.Text = _cees(ks).Dmax2.ToString("F2")    'diameter [mu] 100% catch
        TextBox111.Text = _cees(ks).Dmin2.ToString("F2")    'diameter [mu] 100% loss
        TextBox116.Text = _istep.ToString("F5")             'Calculation step stage #2

        TextBox117.Text = _cees(ks).sum_psd_diff2.ToString("F3")
        TextBox68.Text = _cees(ks).sum_loss2.ToString("F3")
        TextBox69.Text = _cees(ks).sum_loss_C2.ToString("F3")
        TextBox120.Text = Eff_comb.ToString("F3")

        If CheckBox3.Checked Then   'Dust load correction stage #2
            TextBox65.Text = _cees(ks).loss_total2.ToString("F3")    '[%] Corrected
        Else
            TextBox65.Text = _cees(ks).sum_loss2.ToString("F3")      '[%] NOT Corrected
        End If

        TextBox66.Text = _cees(ks).Efficiency2.ToString("F3")       '[%] stage #2
        TextBox109.Text = _cees(ks).Efficiency2.ToString("F3")      '[%]

        TextBox62.Text = _cees(ks).emmis2_Am3.ToString("F4")        '[gram/Am3] emmissie stage #2
        TextBox108.Text = _cees(ks).emmis2_Am3.ToString("F4")       '[gram/Am3] emmissie stage #2
        TextBox134.Text = _cees(ks).emmis2_Am3.ToString("F4")       '[gram/Am3] emmissie stage #2
        TextBox142.Text = _cees(ks).emmis2_Nm3.ToString("F4")       '[gram/Nm3] emmissie stage #2

        TextBox143.Text = _cees(ks).dust2_Nm3.ToString("F1")    'Dust load [gram/Nm3]
    End Sub
    'Determine the particle diameter class upper and lower limits
    ' Private Function Size_classification(dia As Double, noi As Integer) As Double
    Public Sub Size_classification(ByRef g As GvG_Calc_struct)
        If g.dia > 0 Then
            Select Case True
                Case g.dia < NumericUpDown15.Value      '0-0.3 [mu]
                    g.i_d1 = 0.001                      'Diameter small [mu]
                    g.i_d2 = NumericUpDown15.Value      'Diameter big [mu]
                    g.i_p1 = numericUpDown6.Value / 100 'User lower input percentage
                    g.i_p1 = numericUpDown7.Value / 100 'User upper input percentage
                    g.i_grp = 0
                Case g.dia >= NumericUpDown15.Value And g.dia < NumericUpDown23.Value   '>=10 and < 15
                    g.i_d1 = NumericUpDown15.Value      'Diameter small [mu]
                    g.i_d2 = NumericUpDown23.Value      'Diameter big [mu]
                    g.i_p1 = numericUpDown6.Value / 100 'User lower input percentage
                    g.i_p2 = numericUpDown7.Value / 100 'User upper input percentage
                    g.i_grp = 1
                Case g.dia >= NumericUpDown23.Value And g.dia < NumericUpDown24.Value
                    g.i_d1 = NumericUpDown23.Value      'Diameter small [mu]
                    g.i_d2 = NumericUpDown24.Value      'Diameter big [mu]
                    g.i_p1 = numericUpDown7.Value / 100 'User lower input percentage
                    g.i_p2 = numericUpDown8.Value / 100 'User upper input percentage
                    g.i_grp = 2
                Case g.dia >= NumericUpDown24.Value And g.dia < NumericUpDown25.Value
                    g.i_d1 = NumericUpDown24.Value
                    g.i_d2 = NumericUpDown25.Value
                    g.i_p1 = numericUpDown8.Value / 100
                    g.i_p2 = numericUpDown9.Value / 100
                    g.i_grp = 3
                Case g.dia >= NumericUpDown25.Value And g.dia < NumericUpDown26.Value
                    g.i_d1 = NumericUpDown25.Value
                    g.i_d2 = NumericUpDown26.Value
                    g.i_p1 = numericUpDown9.Value / 100
                    g.i_p2 = numericUpDown10.Value / 100
                    g.i_grp = 4
                Case g.dia >= NumericUpDown26.Value And g.dia < NumericUpDown27.Value
                    g.i_d1 = NumericUpDown26.Value
                    g.i_d2 = NumericUpDown27.Value
                    g.i_p1 = numericUpDown10.Value / 100
                    g.i_p2 = numericUpDown11.Value / 100
                    g.i_grp = 5
                Case g.dia >= NumericUpDown27.Value And g.dia < NumericUpDown28.Value
                    g.i_d1 = NumericUpDown27.Value
                    g.i_d2 = NumericUpDown28.Value
                    g.i_p1 = numericUpDown11.Value / 100
                    g.i_p2 = numericUpDown12.Value / 100
                    g.i_grp = 6
                Case g.dia >= NumericUpDown28.Value And g.dia < NumericUpDown29.Value
                    g.i_d1 = NumericUpDown28.Value
                    g.i_d2 = NumericUpDown29.Value
                    g.i_p1 = numericUpDown12.Value / 100
                    g.i_p2 = numericUpDown13.Value / 100
                    g.i_grp = 7
                Case g.dia >= NumericUpDown29.Value And g.dia < NumericUpDown35.Value
                    g.i_d1 = NumericUpDown29.Value
                    g.i_d2 = NumericUpDown35.Value
                    g.i_p1 = numericUpDown13.Value / 100
                    g.i_p2 = NumericUpDown38.Value / 100
                    g.i_grp = 8
                Case g.dia >= NumericUpDown35.Value And g.dia < NumericUpDown36.Value
                    g.i_d1 = NumericUpDown35.Value
                    g.i_d2 = NumericUpDown36.Value
                    g.i_p1 = NumericUpDown38.Value / 100
                    g.i_p2 = NumericUpDown39.Value / 100
                    g.i_grp = 9
                Case g.dia >= NumericUpDown36.Value And g.dia < NumericUpDown37.Value
                    g.i_d1 = NumericUpDown36.Value
                    g.i_d2 = NumericUpDown37.Value
                    g.i_p1 = NumericUpDown39.Value / 100
                    g.i_p2 = NumericUpDown40.Value / 100
                    g.i_grp = 10
                Case g.dia >= NumericUpDown37.Value And g.dia < NumericUpDown51.Value
                    g.i_d1 = NumericUpDown37.Value          'Diameter small
                    g.i_d2 = NumericUpDown51.Value          'Diameter big
                    g.i_p1 = NumericUpDown40.Value / 100    'User lower input percentage
                    g.i_p2 = NumericUpDown48.Value / 100    'User upper input percentage
                    g.i_grp = 11
                Case Else
                    g.i_d1 = NumericUpDown51.Value          'Diameter small [mu]
                    g.i_d2 = 1000                           'Diameter big [mu]
                    g.i_p1 = NumericUpDown48.Value / 100    'User lower input percentage
                    g.i_p2 = 0                              'User upper input percentage
                    g.i_grp = 12
            End Select

            Dim w(11) As Double  'Individual particle class weights

            w(0) = NumericUpDown48.Value
            w(1) = NumericUpDown40.Value - w(0)
            w(2) = NumericUpDown39.Value - w(1) - w(0)
            w(3) = NumericUpDown38.Value - w(2) - w(1) - w(0)
            w(4) = numericUpDown13.Value - w(3) - w(2) - w(1) - w(0)
            w(5) = numericUpDown12.Value - w(4) - w(3) - w(2) - w(1) - w(0)
            w(6) = numericUpDown11.Value - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(7) = numericUpDown10.Value - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(8) = numericUpDown9.Value - w(7) - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(9) = numericUpDown8.Value - w(8) - w(7) - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(10) = numericUpDown7.Value - w(9) - w(8) - w(7) - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)
            w(11) = numericUpDown6.Value - w(10) - w(9) - w(8) - w(7) - w(6) - w(5) - w(4) - w(3) - w(2) - w(1) - w(0)

            TextBox50.Text = w(11).ToString("0.0")
            TextBox49.Text = w(10).ToString("0.0")
            TextBox46.Text = w(9).ToString("0.0")
            TextBox45.Text = w(8).ToString("0.0")
            TextBox44.Text = w(7).ToString("0.0")
            TextBox43.Text = w(6).ToString("0.0")
            TextBox27.Text = w(5).ToString("0.0")
            TextBox25.Text = w(4).ToString("0.0")
            TextBox81.Text = w(3).ToString("0.0")
            TextBox82.Text = w(2).ToString("0.0")
            TextBox154.Text = w(1).ToString("0.0")

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
            NumericUpDown51.BackColor = CType(IIf(NumericUpDown51.Value >= NumericUpDown36.Value, Color.LightGreen, Color.Red), Color)
        Else
            'MessageBox.Show("Error in line 1946")
        End If
    End Sub
    Private Sub Calc_cycl_weight()
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
        Dim c_area1, c_area2 As Double      'Outside area est.

        '========== Stage cyclone #1 =====
        plt_top1 = NumericUpDown32.Value        '[mm] top plate
        plt_body1 = NumericUpDown31.Value       '[mm] rest of the cyclone
        _db = numericUpDown5.Value / 1000       '[m] Body diameter

        'weight top plate
        w1 = PI / 4 * _db ^ 2 * plt_top1 / 1000 * ro_steel

        'weight cylindrical body
        hh = _cyl1_dim(10) * _db / 1000                     '[m] Length romp
        w2 = PI * _db * hh * plt_body1 * ro_steel           '[kg] weight romp

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
        c_area1 = c_weight1 / (plt_body1 / 1000 * ro_steel) 'Area [m2]


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

        c_area2 = C_weight2 / (plt_body2 / 1000 * ro_steel)  'Area [m2]

        total_instal = c_weight1 * NumericUpDown20.Value
        total_instal += C_weight2 * NumericUpDown33.Value

        sheet_metal_wht = total_instal * 1.45               'Gross Sheet metal weight
        p304 = sheet_metal_wht * NumericUpDown45.Value      'rvs 304
        p316 = sheet_metal_wht * NumericUpDown44.Value      'rvs 316

        '-------- present -----------
        TextBox72.Text = C_weight2.ToString("F0")
        TextBox127.Text = c_area2.ToString("F2")            'Area [m2]
        TextBox99.Text = total_instal.ToString("F0")
        TextBox83.Text = sheet_metal_wht.ToString("F0")
        TextBox114.Text = p304.ToString("F0")
        TextBox115.Text = p316.ToString("F0")
        TextBox61.Text = c_weight1.ToString("F0")           'Total weight
        TextBox135.Text = c_area1.ToString("F2")            'Area [m2]
    End Sub
    'Calculate cyclone weight
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabPage7.Enter, NumericUpDown32.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown42.ValueChanged, NumericUpDown41.ValueChanged
        Calc_cycl_weight()
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Save_present_case_to_array()
    End Sub
    Private Sub Save_present_case_to_array()
        Dim cc As Integer
        'Save data of screen into the _cees array

        cc = CInt(NumericUpDown30.Value)       'Case number
        TextBox24.Text &= "line 2267, save case to array nr" & cc.ToString & vbCrLf
        Fill_array_from_screen(cc)
    End Sub

    Private Sub NumericUpDown30_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown30.ValueChanged
        Dim zz As Integer = CInt(NumericUpDown30.Value)    'Case number

        Fill_Screen_from_array(zz)
    End Sub
    Private Sub Fill_Screen_from_array(zz As Integer)
        Dim p1_rel As Double

        TextBox24.Text &= "line 2276, Fill_Screen_from_array nr" & zz.ToString & vbCrLf

        If _cees(zz).case_name.Length > 0 Then

            '----------- General (not calculated) data------------------
            TextBox28.Text = _cees(zz).Quote_no                 'Quote number
            TextBox29.Text = _cees(zz).Tag_no                   'The Tag number
            TextBox53.Text = _cees(zz).case_name                'Case name

            Chck_value(NumericUpDown1, CDec(_cees(zz).FlowT))       'Air flow total
            Chck_value(NumericUpDown4, CDec(_cees(zz).dust1_Am3))   'Dust inlet [g/Am3] 

            ComboBox1.SelectedIndex = _cees(zz).Ct1                 'Cyclone type stage #1
            ComboBox2.SelectedIndex = _cees(zz).Ct2                 'Cyclone type stage #2

            Chck_value(NumericUpDown20, _cees(zz).Noc1)              'Cyclone in parallel
            Chck_value(numericUpDown13, CDec(_cees(zz).db1))         'Diameter cyclone body #1
            Chck_value(NumericUpDown34, CDec(_cees(zz).db2))         'Diameter cyclone body #2
            Chck_value(numericUpDown3, CDec(_cees(zz).ro_gas))       'Density [kg/hr]
            Chck_value(numericUpDown2, CDec(_cees(zz).ro_solid))     'Density [kg/hr]
            Chck_value(numericUpDown14, CDec(_cees(zz).visco))       'Visco in Centi Poise
            Chck_value(NumericUpDown18, CDec(_cees(zz).Temp))        'Temperature [c]
            p1_rel = (_cees(zz).p1_abs - 101325) / 100
            Chck_value(NumericUpDown19, CDec(p1_rel))                'Pressure [Pa abs]-->[mbar g]

            '[mu] Class upper particle diameter limit diameter
            Chck_value(NumericUpDown15, CDec(_cees(zz).dia_big(0)))   '10
            Chck_value(NumericUpDown23, CDec(_cees(zz).dia_big(1)))   '15
            Chck_value(NumericUpDown24, CDec(_cees(zz).dia_big(2)))   '20
            Chck_value(NumericUpDown25, CDec(_cees(zz).dia_big(3)))   '30
            Chck_value(NumericUpDown26, CDec(_cees(zz).dia_big(4)))   '40
            Chck_value(NumericUpDown27, CDec(_cees(zz).dia_big(5)))   '50
            Chck_value(NumericUpDown28, CDec(_cees(zz).dia_big(6)))   '60
            Chck_value(NumericUpDown29, CDec(_cees(zz).dia_big(7)))   '80

            'addeed
            Chck_value(NumericUpDown35, CDec(_cees(zz).dia_big(8)))   '40
            Chck_value(NumericUpDown36, CDec(_cees(zz).dia_big(9)))   '50
            Chck_value(NumericUpDown37, CDec(_cees(zz).dia_big(10)))   '60
            Chck_value(NumericUpDown51, CDec(_cees(zz).dia_big(11)))   '80


            'Percentage van de inlaat stof belasting
            Chck_value(numericUpDown6, CDec(_cees(zz).class_load(0) * 100))
            Chck_value(numericUpDown7, CDec(_cees(zz).class_load(1) * 100))
            Chck_value(numericUpDown8, CDec(_cees(zz).class_load(2) * 100))
            Chck_value(numericUpDown9, CDec(_cees(zz).class_load(3) * 100))
            Chck_value(numericUpDown10, CDec(_cees(zz).class_load(4) * 100))
            Chck_value(numericUpDown11, CDec(_cees(zz).class_load(5) * 100))
            Chck_value(numericUpDown12, CDec(_cees(zz).class_load(6) * 100))
            Chck_value(numericUpDown13, CDec(_cees(zz).class_load(7) * 100))

            Chck_value(NumericUpDown38, CDec(_cees(zz).class_load(8) * 100))
            Chck_value(NumericUpDown39, CDec(_cees(zz).class_load(9) * 100))
            Chck_value(NumericUpDown40, CDec(_cees(zz).class_load(10) * 100))
            Chck_value(NumericUpDown48, CDec(_cees(zz).class_load(11) * 100))


            Me.Refresh()
        End If

    End Sub
    Private Sub Chck_value(num As NumericUpDown, value As Decimal)
        'Make sure the numericupdown.value is within the min-max value

        Select Case value
            Case > num.Maximum
                num.Value = num.Maximum
                TextBox24.Text &= "NOK, value= " & value.ToString & ",  num= " & num.Name.ToString & vbCrLf
            Case < num.Minimum
                num.Value = num.Minimum
                TextBox24.Text &= "NOK, value= " & value.ToString & ",  num= " & num.Name.ToString & vbCrLf
            Case Else
                'TextBox24.Text &= "OK, value= " & value.ToString & ",  num= " & num.Name.ToString & vbCrLf
                num.Value = value
        End Select
    End Sub


    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Calc_sequence()
    End Sub

    Private Sub Calc_stage1_2_comb()
        Dim ks As Integer = CInt(NumericUpDown30.Value)
        Dim li(11) As Double
        Dim row As Integer
        Dim w19 As Double   'Dust load [kg/Nm3] 1stage
        Dim w20 As Double   'Dust load [kg/Nm3] 2stage

        If _cees(ks).stage1(row).d_ave > 0 Then   'For fast startup

            DataGridView4.ColumnCount = 10
            DataGridView4.Rows.Clear()
            DataGridView4.Rows.Add(111)
            DataGridView4.EnableHeadersVisualStyles = False                         'For backcolor

            DataGridView4.Columns(0).HeaderText = "Dia class [mu]"                  'Chart
            DataGridView4.Columns(1).HeaderText = "In abs [g/Am3]"
            DataGridView4.Columns(2).HeaderText = "Inlet psd diff [%]"
            DataGridView4.Columns(3).HeaderText = "Inlet psd cum (chart)[%]"
            DataGridView4.Columns(4).HeaderText = "Loss1 pds diff [%]"               '
            DataGridView4.Columns(5).HeaderText = "Loss1 (chart) pdscum [%]"        'chart
            DataGridView4.Columns(6).HeaderText = "Loss abs [g/Nm3]"                '
            DataGridView4.Columns(7).HeaderText = "Loss2 pds diff [%]"              '
            DataGridView4.Columns(8).HeaderText = "Loss2 (chart) pdscum [%]"        'chart
            DataGridView4.Columns(9).HeaderText = "Eff 1&2 (chart) [%]"             'chart

            DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            DataGridView4.Columns(0).HeaderCell.Style.BackColor = Color.Yellow      'For chart
            DataGridView4.Columns(3).HeaderCell.Style.BackColor = Color.Yellow      'For chart
            DataGridView4.Columns(5).HeaderCell.Style.BackColor = Color.Yellow      'For chart
            DataGridView4.Columns(8).HeaderCell.Style.BackColor = Color.Yellow      'For chart
            DataGridView4.Columns(9).HeaderCell.Style.BackColor = Color.Yellow      'For chart

            '===== Dust load 1stage [kg/Nm3]  =====
            w19 = _cees(ks).dust1_Nm3 / 1000 'Dust load [gr/Nm3]

            '===== Emissie 1stage [kg/Nm3]  =====
            w20 = _cees(ks).emmis1_Nm3 / 1000  'Emission [kg/Nm3] 1stage

            '========== first line =============
            row = 0
            li(0) = _cees(ks).stage1(row).d_ave                 'Dia aver [mu]"
            li(1) = _cees(ks).stage1(row).psd_dif * 10 * w19    'In abs [g/Nm3]" T76 * W19
            li(2) = 100 * li(1) / _cees(ks).dust1_Nm3             'In psd diff [%]
            li(3) = 100 - li(2)                                 'In psd diff cumm[%]" S76 
            li(4) = _cees(ks).stage2(row).psd_dif               'In psd cum [%]
            li(5) = 100 - _cees(ks).stage2(row).psd_dif         'Loss1 pdscum [%]
            li(6) = _cees(ks).stage2(row).loss_abs_C * 10 * w20 'Loss abs [g/Nm3]
            li(7) = 1000 * li(6) / _cees(ks).sum_loss_C2        'Loss2 pds dif [%]
            li(8) = 100 - li(7)                                 'Loss2 pds.cum [%]
            li(9) = 100 * (li(1) - li(6)) / li(1)               'Eff stage1&2 [%]

            For col = 0 To 9  'Fill the DataGrid
                If Double.IsNaN(li(col)) Then li(col) = 0           'prevent silly results
                If li(col) < 0 Or li(col) > 10 ^ 5 Then li(col) = 0   'prevent silly results
                DataGridView4.Rows(row).Cells(col).Value = li(col).ToString("F5")
            Next

            '========- rest of the lines ========
            Dim qq As Double
            For row = 1 To DataGridView4.Rows.Count - 1                 'Fill the DataGrid
                li(0) = _cees(ks).stage1(row).d_ave                     'Dia aver [mu]
                li(1) = _cees(ks).stage1(row).psd_dif * 10 * w19        'In abs [g/Am3]" T76 * W19
                li(2) = 100 * li(1) / _cees(ks).dust1_Nm3               'In psd diff [%]
                li(3) = _cees(ks).stage1(row - 1).psd_cum_pro - li(2)   'In psd diff cumm[%]
                li(4) = _cees(ks).stage2(row).psd_dif                   'In psd cum [%]

                qq = CDbl(DataGridView4.Rows(row - 1).Cells(5).Value)
                li(5) = qq - _cees(ks).stage2(row).psd_dif              'Loss1 pdscum [%] 
                li(6) = _cees(ks).stage2(row).loss_abs_C * 10 * w20     'Loss abs [g/Am3]
                li(7) = 1000 * li(6) / _cees(ks).sum_loss_C2            'Loss2 pds dif [%]

                qq = CDbl(DataGridView4.Rows(row - 1).Cells(8).Value)
                li(8) = qq - li(7)                                      'Loss2 pds.cum [%] 
                li(9) = 100 * ((li(1) - li(6)) / li(1))                 'Eff stage1&2 [%]

                If li(8) = 0 Then li(9) = 100       'Loss is zero eff must be 100%

                '========== prevent silly results ======
                For col = 0 To 9  'Fill the DataGridview
                    If Double.IsNaN(li(col)) Then li(col) = 0               'prevent silly results
                    If li(col) < 0 Or li(col) > 10 ^ 5 Then li(col) = 0     'prevent silly results
                    DataGridView4.Rows(row).Cells(col).Value = li(col).ToString("F6")
                Next
            Next

            '---------------- present dp(x) values ----
            TextBox119.Text = Vlookup_db(100).ToString("F2")
            TextBox121.Text = Vlookup_db(95).ToString("F2")
            TextBox122.Text = Vlookup_db(90).ToString("F2")
            TextBox123.Text = Vlookup_db(50).ToString("F2")
            TextBox124.Text = Vlookup_db(10).ToString("F2")
            TextBox125.Text = Vlookup_db(5).ToString("F2")
        End If
    End Sub
    Private Function Vlookup_db(loss As Double) As Double
        'If the loss percenatege is found return the particle diameter
        If loss > 100 Or loss < 0 Then MsgBox("Problem in line Vlookup_db")

        loss = 100 - loss
        For row = 1 To DataGridView4.Rows.Count - 1
            If loss <= CDbl(DataGridView4.Rows(row).Cells(9).Value) Then
                Return CDbl((DataGridView4.Rows(row).Cells(0).Value))
                Exit For
            End If
        Next

        Return (-1)
    End Function

    'Calculate ACTUAL --> NORMAL Conditions
    'Normaal condities; 0 celsius, 101325 Pascal
    'http://www.installbasis.nl/downloads/Omrekening%20Normaalkubiekemeters.PDF
    'PV=MRT     ===>    M/V=ρ  ===>    ρ=P/(R.T)
    'R= P/(ρ.T)
    'R1=R2  ===>   Pn/(ρn.Tn)= P2/(ρ2.T2)
    'Pn.ρ2.T2= P2.ρn.Tn
    'ρn =ρ2.(Pn/P2).(T2/Tn)
    'ρn =ρ2.(101325/P2).(T2/273.15)

    Private Function Calc_Normal_density(ro1 As Double, p1 As Double, t1 As Double) As Double
        Dim ro_normal As Double
        ro_normal = ro1 * (101325 / p1) * ((t1 + 273.15) / 273.15)
        Return (ro_normal)
    End Function

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click, TabPage5.Enter, NumericUpDown47.ValueChanged, NumericUpDown46.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown50.ValueChanged, NumericUpDown49.ValueChanged, NumericUpDown17.ValueChanged
        'Stress calculation
        Calc_Round_plate()
        Calc_shell()
    End Sub
    Private Sub Calc_Round_plate()
        'Round plate simply supported
        Dim dia, r, t As Double
        Dim Elas, p, σm, yt, v As Double
        Dim _fs, sf, _σ_02 As Double
        Dim temper As Double

        _σ_02 = NumericUpDown17.Value           '[N/mm2]
        sf = NumericUpDown49.Value              '[-] safety factor
        _fs = _σ_02 / sf

        p = NumericUpDown16.Value * 100         '[mbar->[N/m2]]

        dia = NumericUpDown47.Value / 1000      '[m]
        r = dia / 2                             '[m]
        t = NumericUpDown46.Value / 1000        '[m]
        temper = NumericUpDown18.Value          '[celsius]

        Elas = (201.66 - 8.48 * temper / 10 ^ 2) '* 10 ^ 9   '[GPa] for 304
        Elas *= 10 ^ 9                                       '[Pa] for 304

        v = 0.3 'For steel
        σm = 1.238 * p * r ^ 2 / t ^ 2
        σm /= 10 ^ 6                                        '[N/mm2]

        yt = (0.696 * p * r ^ 4) / (Elas * t ^ 3)           '[m]
        yt *= 1000                                          '[m]--->[mm]

        TextBox144.Text = temper.ToString("F0")             '[celsius]
        TextBox138.Text = (Elas / 10 ^ 9).ToString("F0")    '[GPa]
        TextBox136.Text = σm.ToString("F0")                 '[N/mm2]
        TextBox137.Text = yt.ToString("F1")                 '[mm]

        '===== check ================
        TextBox136.BackColor = CType(IIf(σm > _fs, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Calc_shell()
        'EN13445, equation 7.4.2 Required wall thickness
        'For symbols, quanties and units see page 12
        Dim dia As Double
        Dim p, e_wall As Double
        Dim _σ_02 As Double
        Dim _fs As Double
        Dim z_joint As Double      'weld joint efficiency
        Dim sf As Double            'Safety factor

        _σ_02 = NumericUpDown17.Value           '[N/mm2]
        sf = NumericUpDown49.Value              '[-] safety factor
        z_joint = NumericUpDown50.Value         '[-] weld factor
        _fs = _σ_02 / sf

        p = NumericUpDown16.Value / 10000           '[mbar->[N/mm2]=[MPa]]
        dia = NumericUpDown47.Value                 '[mm]
        e_wall = p * dia / (2 * _fs * z_joint + p)  'equation (7.4.2) page 30, Required wall thickness

        '---------------------
        TextBox139.Text = p.ToString("F3")          '[N/mm2] pressure
        TextBox140.Text = e_wall.ToString("F1")     '[mm] wall thickness
        TextBox141.Text = _fs.ToString("F0")        '[N/mm2] design stress
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click, TabPage10.Enter
        Draw_chart3(Chart3)             'Dust load correction
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Form2.Show()    'Fopspeen
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click, TabPage13.Enter
        Dim words() As String
        DataGridView5.ColumnCount = 4
        DataGridView5.Rows.Clear()
        DataGridView5.Rows.Add(DSM_psd_example.Length - 1)
        DataGridView5.EnableHeadersVisualStyles = False           'For backcolor

        DataGridView5.Columns(0).HeaderText = "Dia [um]"
        DataGridView5.Columns(1).HeaderText = "Cumm [%]"
        DataGridView5.Columns(2).HeaderText = "--"

        For row = 0 To DataGridView5.Rows.Count - 1
            words = DSM_psd_example(row).Split(CType(";", Char()))
            DataGridView5.Rows(row).Cells(0).Value = CDbl(words(0))
            DataGridView5.Rows(row).Cells(1).Value = CDbl(words(1))
            DataGridView5.Rows(row).Cells(2).Value = (100 - CDbl(words(1)))
        Next

    End Sub

End Class
