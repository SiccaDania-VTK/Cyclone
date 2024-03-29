﻿Imports System.Globalization
Imports System.IO
Imports System.Management
Imports System.Math
Imports System.Text
Imports System.Threading
Imports System.Windows.Forms.DataVisualization.Charting
Imports MathNet
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

    '====== INPUT DATA ======
    Public FlowT As Double          '[Am3/h] Air flow 
    Public ro_gas As Double         '[kg/m3] Density gas
    Public ro_solid As Double       '[kg/m3] Density 
    Public visco As Double          '[Centi Poise] Visco in 
    Public Temp As Double           '[c] Temperature 
    Public Spare1 As String         'For future use
    Public Spare2 As String         'For future use
    Public Spare3 As String         'For future use

    '===== stage #1 parameter ======
    Public Flow1 As Double          '[Am3/s] Air flow per cyclone 
    Public Flow_air_kgh As Double   '[kg/h] Air flow total
    Public dust1_in_kgh As Double   '[kg/h] Dust inlet stage 1 total  
    Public dust1_Am3 As Double      '[g/Am3] Dust load inlet 
    Public dust1_Nm3 As Double      '[g/Nm3] Dust load inlet 
    Public emmis1_Am3 As Double     '[g/Am3] Dust emission 
    Public emmis1_Nm3 As Double     '[g/Nm3] Dust emission 
    Public emis1_kgh As Double      '[kg/h]  Dust emission 
    Public Efficiency1 As Double    '[%] Efficiency Stage #1 
    Public sum_loss1 As Double      '[-]Passed trough cyclone 
    Public sum_loss_C1 As Double    '[-] Passed trough cyclone Corrected
    Public loss_total1 As Double
    Public sum_psd_diff1 As Double
    Public p1_abs As Double         '[Pa abs] pressure inlet abs
    Public dpgas1 As Double         '[Pa] pressure loss gas
    Public dpdust1 As Double        '[Pa] pressure loss dust
    Public Ro_gas1_Am3 As Double    '[kg/Am3]  density gas (inlet stage #1)
    Public Ro_gas1_Nm3 As Double    '[kg/Nm3] density gas (inlet stage #1)
    Public Ct1 As Integer           '[-] Cyclone type (eg AC435)
    Public Noc1 As Double           '[-] Number paralle Cyclones
    Public db1 As Double            '[m] Diameter cyclone body
    Public inh1 As Double           '[m] inlet hoogte
    Public inb1 As Double           '[m] inlet breedte
    Public dout1 As Double          '[m] diameter zakbuis
    Public inv1 As Double           '[m/s] Inlet velocity cyclone
    Public outv1 As Double          '[m/s] Outlet velocity cyclone
    Public Kstokes1 As Double       'Stokes getal
    Public m1 As Double             'm factor loss curve d< dia critical
    Public stage1 As GvG_Calc_struct()   'tbv calculatie stage #1
    Public Dmin1 As Double          'Smallest particle 100% loss
    Public Dmax1 As Double          'Biggest particle 100% catch

    '===== stage #2 parameters ======
    Public Flow2 As Double          '[Am3/s] Air flow per cyclone 
    Public dust2_in_kgh As Double   '[kg/h] Dust inlet stage 1 total 
    Public dust2_Am3 As Double      '[g/Am3] Dust load inlet (inlet stage #2)
    Public dust2_Nm3 As Double      '[g/Nm3] Dust load inlet (inlet stage #2) 
    Public emmis2_Am3 As Double     '[g/Am3] Dust emission 
    Public emmis2_Nm3 As Double     '[g/Nm3] Dust emission 
    Public emis2_kgh As Double      '[kg/h]  Dust emission 
    Public sum_loss2 As Double      'Passed trough cyclone
    Public sum_loss_C2 As Double    'Passed trough cyclone Corrected
    Public loss_total2 As Double
    Public sum_psd_diff2 As Double
    Public Efficiency2 As Double    'Efficiency Stage #1 [%}
    Public p2_abs As Double         '[Pa abs] pressure inlet abs (inlet 2nd stage)
    Public dpgas2 As Double         '[Pa] pressure loss gas (inlet 2nd stage)
    Public dpdust2 As Double        '[Pa] pressure loss dust
    Public Ro_gas2_Am3 As Double    '[kg/Am3] density gas
    Public Ro_gas2_Nm3 As Double    '[kg/Nm3] density gas
    Public Ct2 As Integer           '[-] Cyclone type (eg AC435)
    Public Noc2 As Double           '[-] Number paralle Cyclones
    Public db2 As Double            '[m] Diameter cyclone body
    Public inh2 As Double           '[m] inlet hoogte
    Public inb2 As Double           '[m] inlet breedte
    Public dout2 As Double          '[m] diameter zakbuis
    Public inv2 As Double           '[m/s] Inlet velocity cyclone
    Public outv2 As Double          '[m/s] Outlet velocity cyclone
    Public Kstokes2 As Double       'Stokes of the particle

    '===== stage #3 parameters (Outlet stage2) ======
    Public p3_abs As Double         '[Pa abs] pressure outlet abs (outlet stage #2)
    Public Ro_gas3_Am3 As Double    '[kg/Am3] density gas (outlet stage #2)
    Public Ro_gas3_Nm3 As Double    '[kg/Nm3] density gas (outlet stage #2)

    Public m2 As Double             'm factor loss curve d< dia critical
    Public stage2 As GvG_Calc_struct()   'tbv calculatie stage #2
    Public Dmin2 As Double          '[mu] Smallest particle 100% loss
    Public Dmax2 As Double          '[mu] Biggest particle 100% catch
End Structure

'Variables used by GvG in calculation
<Serializable()> Public Structure GvG_Calc_struct
    Public dia As Double            '[mu] Particle diameter 
    Public d_ave As Double          '[mu] Average diameter 
    Public d_ave_K As Double        '[-] Average diam/K_stokes 
    Public i_grp As Double          'Particle Groepnummer (stage 2= stage 1)
    Public i_d1 As Double           '[mu] smallest particle diameter in Class
    Public i_d2 As Double           '[mu] biggest particle diameter in Class
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
'Variables used by GvG in calculation
<Serializable()> Public Structure Psd_input_struct
    Public dia_big As Double      '[mu] Particle diameter inlet cyclone (input data)
    Public class_load As Double   '[% weight] group_weight_cum in de inlaat stroom 
End Structure

Public Class Form1
    Public Const no_PDS_inputs As Integer = 300         'was 100, Data from the customer
    Private Const V As Boolean = False
    Private Const V1 As Boolean = False
    Public _cyl1_dim(20) As Double                      'Cyclone stage #1 dimensions
    Public _cyl2_dim(20) As Double                      'Cyclone stage #2 dimensions
    Public _istep As Double                             '[mu] Particle size calculation step
    Public _cees(20) As Input_struct                    '20 Case's data     (case 0 is for calculations ONLY)
    Public _input(no_PDS_inputs + 1) As Psd_input_struct    'Particle Size Distribution
    Private k41 As Double                                   'sum loss abs (for DataGridView1)
    Private init As Boolean = False                         'Initialize done
    Private Update_screen_fram_array_done As Boolean = True 'Program chokes during retrieve
    Private ReadOnly _Gasconstant As Double = 8.31445984848     'ideal gas constant 
    Private ReadOnly separators As String() = {";"}

    Public _ν As Double = 0.3       'Poisson ratio for steel
    Public _P As Double             'Calculation pressure [Mpa]
    Public _fs As Double = 1        'Allowable stress shell [N/mm2]
    Public _f02 As Double = 1       'Yield 0.2% stress shell [N/mm2]
    Public _fym As Double = 1       'Allowable stress reinforcement [N/mm2]

    Public _De As Double            'Outside diameter shell
    Public _Di As Double            'Inside diameter shell
    Public _ecs As Double           'Shell thickness

    Public _deb As Double           'Outside diameter nozzle fitted in shell
    Public _dib As Double           'Inside diameter nozzle fitted in shell
    Public _Emod As Double          'Modulus of elasticity 

    Public Shared joint_eff As String() = {"0.7", "0.85", "1.0"}    'Welding

    '===== Testing with Normal-log distribution ======
    Public norm_log_dist_pdf(300, 2) As Double   'Normal-log distribution, Bell Curve
    Public norm_log_dist_cdf(300, 2) As Double   'Normal-log distribution, Cumulative wght

    'Type AC;Inlaatbreedte;Inlaathoogte;Inlaatlengte;Inlaat hartmaat;Inlaat afschuining;
    'Uitlaat keeldia inw.;Uitlaat flensdiameter inw.;Lengte insteekpijp inw.;
    'Lengte romp + conus;Lengte romp;Lengte conus;Dia_conus / 3P-pijp;Lengte 3P-pijp;Lengte 3P-conus;Kleine dia 3P-conus",

    Private ReadOnly cyl_dimensions As String() = {
    "AC-300;0.34        ;0.770;0.600;0.630;0.300;0.680  ;0.6956;0.892;3.360;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-350;0.32        ;0.700;0.600;0.617;0.300;0.630  ;0.6456;0.892;3.360;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-435;0.282       ;0.640;0.600;0.600;0.300;0.560  ;0.5756;0.892;3.360;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-550;0.25        ;0.570;0.600;0.580;0.300;0.450  ;0.5600;0.892;3.360;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-750;0.216       ;0.486;0.600;0.570;0.300;0.365  ;0.5600;0.892;3.360;1.312;2.048;0.4;0.6;0.6;0.25",
    "AC-850;0.203       ;0.457;0.600;0.564;0.300;0.307  ;0.4280;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-1850;0.136      ;0.310;0.600;0.530;0.300;0.150  ;0.2500;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-850+afz;0.203   ;0.457;0.600;0.564;0.300;0.307  ;0.4280;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AC-1850+afz;0.136  ;0.310;0.600;0.530;0.300;0.150  ;0.2500;0.892;3.797;1.312;2.485;0.4;0.6;0.6;0.25",
    "AA850;0.203;0.457  ;0.600;0.564;0.300;0.307;0.428  ;0.8920;3.797;1.312;2.485;0.4;0.6;0.6;0.25"}    'NOT UP TO DATE CHECK !!!!!!
    '"Void;0.203;0.457  ;0.600;0.564;0.300;0.307;0.428  ;0.8920;3.797;1.312;2.485;0.4;0.6;0.6;0.25"     'NOT UP TO DATE CHECK !!!!!!
    '}

    Private ReadOnly Tangent_out_dimensions As String() = {
    "AC-300 ;0.6956;0.5440;0.5440   ;0.5100;0.3400;0.0340;0.5780;0.0680",
    "AC-350 ;0.6456;0.5040;0.5040   ;0.4725;0.3150;0.0315;0.5355;0.0630",
    "AC-435 ;0.5756;0.4480;0.4480   ;0.4200;0.2800;0.0280;0.4760;0.0560",
    "AC-550 ;0.5600;0.4480;0.4480   ;0.4200;0,2800;0.0280;0.4760;0.0560",
    "AC-750 ;0.5600;0.4480;0.4480   ;0.4200;0.2800;0.0280;0.4760;0.0560",
    "AC-850 ;0.4280;0.3424;0.3424   ;0.3210;0.2140;0.0214;0.3638;0.0428",
    "AC-1850;0.2500;0.2000;0.2000   ;0.1875;0.1250;0.0125;0.2125;0.0250"
    }

    'Nieuwe reken methode, verdeling volgens Weibull verdeling
    'm1,k1,a1 als d < d_krit
    'm2,k2,a2 als d > d_krit
    'type; d/krit; m1; k1; a1; m2; k2; a2; drukcoef air; drukcoef dust
    Private ReadOnly rekenlijnen As String() = {
    "AC300;     12.2;   1.15;   7.457;  1.005;      8.5308;     1.6102; 0.4789; 7;      0",
    "AC350;     10.2;   1.0;    5.3515; 1.0474;     4.4862;     2.4257; 0.6472; 7;      7.927",
    "AC435.;    8.93;   0.69;   4.344;  1.139;      4.2902;     1.3452; 0.5890; 7;      8.26",
    "AC550;     8.62;   0.527;  3.4708; 0.9163;     3.3211;     1.7857; 0.7104; 7;      7.615",
    "AC750;     8.3;    0.50;   2.8803; 0.8355;     4.0940;     1.0519; 0.6010; 7.5;    6.606",
    "AC850.;    7.8;    0.52;   1.9418; 0.73705;    -0.1060;    2.0197; 0.7077; 9.5;    6.172",
    "AC850+afz; 10;     0.5187; 1.6412; 0.8386;     4.2781;     0.06777;0.3315; 0;      0",
    "AC1850;    9.3;    0.50;   1.1927; 0.5983;     -0.196;     1.3687; 0.6173; 14.5;   0",
    "AC1850+afz;10.45;  0.4617; 0.2921; 0.4560;     -0.2396;    0.1269; 0.3633; 0;      0",
    "AA850;     7.8;    0.52;   1.9418; 0.73705;    -0.1060;    2.0197; 0.7077; 10;     6.8",
    "Void;      7.8;    0.52;   1.9418; 0.73705;    -0.1060;    2.0197; 0.7077; 0;      0"
    }

    'DSM polymer power DSM Geleen (cumulatief[%], particle diameter[mu])
    Private ReadOnly DSM_psd_example As String() = {
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
    "215;99.95"}

    'Typical Corn ( particle diameter[mu],cumulatief[%] )
    Private ReadOnly psd_corn As String() = {
    "1;	99.99",
    "2;	99.9",
    "4;	99.8",
    "6;	95",
    "8;	84",
    "10;70",
    "12;52.2",
    "14;35",
    "16;22.6",
    "18;15",
    "20;10",
    "24;6.5",
    "28;3",
    "32;1",
    "36;0.5",
    "40;0.1",
    "44;0.01",
    "48;0.001"}

    'Typical Chickpea Starch (particle diameter[mu], cumulatief[%])
    Private ReadOnly psd_chickpea_starch As String() = {
    "6.1590;	99.9998",
    "6.7610;	99.995",
    "7.4000;	99.96",
    "8.2000;	99.83",
    "9.0000;	99.5",
    "9.8000;	98.86",
    "10.8000;	97.77",
    "11.8000;	96.06",
    "13.0000;	93.54",
    "14.2000;	89.99",
    "15.7000;	85.25",
    "17.2000;	79.21",
    "18.9000;	71.91",
    "20.7000;	63.54",
    "22.7000;	54.51",
    "24.9500;	45.42",
    "27.4000;	36.93",
    "30.1000;	29.63",
    "33.0000;	23.85",
    "36.2000;	19.59",
    "39.8000;	16.57",
    "43.7000;	14.37",
    "47.9000;	12.58",
    "52.6000;	10.92",
    "57.8000;	9.26",
    "63.2000;	7.61",
    "69.6000;	6.07",
    "76.4000;	4.75",
    "83.9000;	3.70",
    "92.1000;	2.90",
    "101.0000;	2.29",
    "111.0000;	1.78",
    "121.8000;	1.33",
    "133.7000;	0.92",
    "146.8000;	0.56",
    "161.2000;	0.29",
    "176.9000;	0.12",
    "194.2000;	0.046",
    "213.2000;	0.026",
    "234.0000;	0.0234",
    "256.9000;	0.023316"}


    'SURESH, Cargill China (particle diameter[mu],cumulatief[%] )
    Private ReadOnly maltodextrine_psd_suresh As String() = {
    "1.0	;	99.9999",
    "3.5	;	97.00",
    "7.5	;	93.80",
    "12.5	;	86.10",
    "17.5	;	75.40",
    "25.0	;	66.00",
    "35.0	;	51.10",
    "45.0	;	39.70",
    "55.0	;	29.40",
    "65.0	;	23.50",
    "75.0	;	20.90",
    "85.0	;	18.50",
    "95.0	;	13.80",
    "110.0	;	11.70",
    "140.0	;	7.90",
    "180.0	;	3.50",
    "225.0	;	0.80",
    "275.0	;	0.10",
    "350.0	;	0.01"}

    'Cargill China (particle diameter[mu], cumulatief[%])
    Private ReadOnly maltodextrine_psd As String() = {
    "0.359;99.999	",
    "0.652;99.990	",
    "0.717;99.980	",
    "0.799;99.940	",
    "0.899;99.900	",
    "1.001;99.840	",
    "1.116;99.770	",
    "1.242;99.700	",
    "1.379;99.610	",
    "1.535;99.510	",
    "1.709;99.400	",
    "1.902;99.270	",
    "2.117;99.120	",
    "2.356;98.920	",
    "2.622;98.740	",
    "2.919;98.510	",
    "3.249;98.240	",
    "3.611;97.950	",
    "4.020;97.620	",
    "4.478;97.250	",
    "4.984;96.850	",
    "5.549;96.390	",
    "6.175;95.890	",
    "6.873;95.330	",
    "7.650;94.700	",
    "8.515;94.000	",
    "9.435;93.200	",
    "10.505;92.300	",
    "11.745;91.290	",
    "13.070;90.160	",
    "14.540;88.920	",
    "16.410;87.590	",
    "18.235;86.180	",
    "20.090;84.680	",
    "22.360;83.140	",
    "24.835;81.520	",
    "27.640;79.810	",
    "30.630;77.960	",
    "34.605;75.930	",
    "38.610;73.650	",
    "42.520;71.060	",
    "47.315;68.200	",
    "52.550;64.720	",
    "58.485;60.870	",
    "65.095;56.560	",
    "70.955;51.820	",
    "79.140;46.720	",
    "89.750;41.350	",
    "99.875;35.840	",
    "111.150;30.0 ",
    "140.0	;   20",
    "180.0	;	10",
    "225.0	;	4",
    "275.0	;	1",
    "350.0	;	0.1",
    "400.0	;	0.01",
    "450.0	;	0.001"}

    'GvG Excelsheet, Cumulatief[%], particle diameter[mu])
    Private ReadOnly GvG_excel As String() = {
    "10;   99.2",
    "15;   85.2",
    "20;   57.5",
    "30;   15.4",
    "40;   10",
    "50;   6.7",
    "60;   4.5",
    "80;   2",
    "100;  0"}

    'AA850 test GvG Excelsheet, Cumulatief[%], particle diameter[mu])
    Private ReadOnly AA_excel As String() = {
    "2;   99.9973",
    "3;   99.92",
    "4;   99.61",
    "6;   97.3",
    "8;   91.72",
    "11;   76.4",
    "14;   53.1",
    "16;   35.5",
    "18;  19.5",
    "20;   8.7",
    "22;   3.3",
    "50;  0.001"}

    'Whey (cumulatief[%], particle diameter[mu])
    Private ReadOnly psd_whey_A6605 As String() = {
    "3.0;	99.975",
    "5.0;	99.955",
    "8.0;	99.92",
    "15;	99.8",
    "25;	99.5",
    "50;	98",
    "100;	88",
    "200;	58",
    "400;	20",
    "600;	8",
    "1000;	2",
    "1200;	0.1"}

    'Potato from Flash drier (cumulatief[%], particle diameter[mu])
    Private ReadOnly psd_potato_flash_drier As String() = {
    "5;	    99.999",
    "7;	    99.7",
    "10;	99.0",
    "15;	97.0",
    "20;	91.0",
    "30;	71",
    "50;	32",
    "70;	10",
    "90;	2",
    "100;	1",
    "110;	0.1",
    "120;	0.01"}

    'EN 10028-2 for steel
    'EN 10028-7 for stainless steel
    Public Shared steel As String() = {
   "Material-------;50c;100;150;200;250;300;350;400;450;500;550;remarks--;cs/ss",
   "1.0425 (P265GH);265;241;223;205;188;173;160;150;  0;  0;  0; max 400c;cs",
   "1.0473 (P355GH);343;323;299;275;252;232;214;202;  0;  0;  0; max 400c;cs",
   "1.4301 (304)   ;190;157;142;127;118;110;104; 98; 95; 92; 90; max 550c;ss",
   "1.4307 (304L)  ;180;147;132;118;108;100; 94; 89; 85; 81; 80; max 550c;ss",
   "1.4401 (316)   ;204;177;162;147;137;127;120;115;112;110;108; max 550c;ss",
   "1.4404 (316L)  ;200;166;152;137;127;118;113;108;103;100; 98; max 550c;ss"}

    'Typical Lognatural distribution; Weibull shape; Weibull scale
    Public Shared typ_distri As String() = {
    "Free selection;0;0",
    "Corn;3.4;14.0",
    "Chickpea;4.0;25.0",
    "Potato;2.4;47",
    "Maltodex;1.0;121",
    "Hi protein Pet food powder;2.1;72.4"}


    'Chapter 6, Max allowed values for pressure parts
    Public Shared chap6 As String() = {
   "Chap 6.2, Steel, safety, rupture < 30%; 1.5",
   "Chap 6.4, Austenitic steel, rupture 30-35%; 1.5",
   "Chap 6.5, Austenitic steel, rupture >35%; 3.0",
   "Chap 6.6, Cast steel; 1.9"}

    '----------- directory's-----------
    Private ReadOnly dirpath_Eng As String = "N:\Engineering\VBasic\Cyclone_sizing_input\"
    Private ReadOnly dirpath_Rap As String = "N:\Engineering\VBasic\Cyclone_rapport_copy\"
    Private ReadOnly dirpath_tmp As String = "C:\Tmp\"
    Private ReadOnly ProcID As Integer = Process.GetCurrentProcess.Id
    Private ReadOnly dirpath_Temp As String = "C:\Temp\" & ProcID.ToString

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim hh, life_time, i As Integer
        Dim words As String()
        Dim separators As String() = {";"}
        Dim Pro_user, HD_number As String
        Dim nu, nu2 As Date
        Dim user_list As New List(Of String)
        Dim hard_disk_list As New List(Of String)
        Dim pass_name As Boolean = False
        Dim pass_disc As Boolean = False

        With DataGridView1
            .ColumnCount = 9
            .Rows.Clear()
            .Rows.Add(23)
            .RowHeadersVisible = False
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        End With

        'Initialize the arrays in the struct
        For i = 0 To _cees.Length - 1
            _cees(i).case_name = ""
            ReDim _cees(i).stage1(111)                        'Initialize
            ReDim _cees(i).stage2(111)                        'Initialize
        Next

        '------ allowed users with hard disc id's -----
        user_list.Add("GP")
        user_list.Add("gpath")
        user_list.Add("GerritP")
        user_list.Add("gerrit.pathuis")
        user_list.Add("user")
        hard_disk_list.Add("058F63646471")          'Privee PC, graslaan25
        hard_disk_list.Add("50026B768223EE72")      'Desktop Privee PC, graslaan25
        hard_disk_list.Add("S5RRNF0R784528M")       'VTK PC, GP
        hard_disk_list.Add("S2R6NX0H740154H")       'VTK PC, GP
        hard_disk_list.Add("0008_0D02_003E_0FBB.")  'VTK laptop, GP
        hard_disk_list.Add("S28ZNXAG521979")        'VTK laptop, GP privee

        user_list.Add("FredKo")
        user_list.Add("fred.korbeeck")
        hard_disk_list.Add("UGXVK01J1BBS7O")  'VTK new laptop, FKo

        user_list.Add("JanK")
        hard_disk_list.Add("0025_38B4_71B4_88FC.")  'VTK laptop, Jank

        user_list.Add("aj.van.gelder")
        hard_disk_list.Add("0008_0D02_003D_27DB.")  'VTK laptop, AJ

        user_list.Add("Suresh.Sundararaj")
        hard_disk_list.Add("5CD2_E42A_81A1_0D0E.")  'Suresh.Sundararaj Denmark
        hard_disk_list.Add("4C530000100828121514")  'Suresh.Sundararaj Denmark

        user_list.Add("beko")                       '02/11/2021
        user_list.Add("beji.kort")                  '01/03/2022
        hard_disk_list.Add("0025_38AA_01B5_BD94.")  'VTK laptop, Benji Kort


        user_list.Add("jeroen.agricola")
        hard_disk_list.Add("171095402070")          'VTK desktop, Jeroen
        hard_disk_list.Add("170228801578")          'VTK laptop, Jeroen disk 1
        hard_disk_list.Add("MCDBM1M4F3QRBEH6")      'VTK laptop, Jeroen disk 2
        hard_disk_list.Add("0025_388A_81BB_14B5.")  'Zweet kamer, Jeroen 

        user_list.Add("Peterdw")
        user_list.Add("peter.de.wildt")             '01/03/2022
        hard_disk_list.Add("134309552747")          'VTK PC, Peter de Wild

        user_list.Add("bert.korbeeck")
        hard_disk_list.Add("0025_3886_01E9_11D6.")  'VTK new desktop. BKo (24/11/2020)
        hard_disk_list.Add("NA8QWR8W")              'VTK new intallatie BKo (23/02/2022)

        nu = Now()
        nu2 = CDate("2022-06-01 00:00:00")
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
            If HD_number = Trim(hard_disk_list(i)) Then pass_disc = True
        Next

        If pass_name = V OrElse pass_disc = V Then
            Clipboard.SetText("User= " & Pro_user & ", HD=  " & HD_number)
            MessageBox.Show("VTK Cyclone selection program" & vbCrLf & "Access denied, contact GPa" & vbCrLf)
            MessageBox.Show("User_name= " & Pro_user & ", Pass name= " & pass_name.ToString)
            MessageBox.Show("HD_id= *" & HD_number & "*" & ", Pass disc= " & pass_disc.ToString)
            MessageBox.Show("Name and number are copied to the Clipboard !")
            Environment.Exit(0)
        End If

        If life_time < 0 Then
            MessageBox.Show("Program lease Is Expired, contact GPa")
            Environment.Exit(0)
        End If

        Rights_Control()        'Indicate the tab visible to the user

        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        For hh = 0 To (cyl_dimensions.Length - 1)  'Fill combobox1 cyclone types
            words = cyl_dimensions(hh).Split(CType(";", Char()))
            ComboBox1.Items.Add(words(0))
            ComboBox2.Items.Add(words(0))
        Next hh

        ComboBox3.Items.Clear()
        For hh = 0 To (joint_eff.Length - 1)  'Fill combobox 
            ComboBox3.Items.Add(joint_eff(hh))
        Next hh

        ComboBox4.Items.Clear()
        For hh = 0 To (chap6.Length - 1)  'Fill combobox 
            words = chap6(hh).Split(separators, StringSplitOptions.None)
            ComboBox4.Items.Add(words(0))
        Next hh

        ComboBox5.Items.Clear()
        For hh = 1 To (steel.Length - 1)  'Fill combobox steel
            words = steel(hh).Split(separators, StringSplitOptions.None)
            ComboBox5.Items.Add(words(0))
        Next hh

        ComboBox6.Items.Clear()
        For hh = 0 To (typ_distri.Length - 1)  'Fill combobox distribution
            words = typ_distri(hh).Split(separators, StringSplitOptions.None)
            ComboBox6.Items.Add(words(0))
        Next hh

        ComboBox1.SelectedIndex = 2                 'Select Cyclone type AC_435
        ComboBox2.SelectedIndex = 5                 'Select Cyclone type AC_850
        ComboBox3.SelectedIndex = 0                 'Weld joint efficiency
        ComboBox4.SelectedIndex = 1                 'chapter6
        ComboBox5.SelectedIndex = 3                 'Steel 304L
        ComboBox6.SelectedIndex = 0                 'Distribution

        TextBox148.Text = "Het d50 getal geeft de diameter aan waarbij 50% verloren gat 50% wordt gevangen." & vbCrLf
        TextBox148.Text &= "Het d100 getal geeft de diameter aan waarbij 100% verloren gaat." & vbCrLf
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
        TextBox148.Text &= "Siccadania Denmark uses Mastersizer 2000 for PSD determination." & vbCrLf

        TextBox20.Text = "All AA cyclones have a diameter of 300mm" & vbCrLf
        TextBox20.Text &= "Select AC850 and select diameter is 300 mm" & vbCrLf
        TextBox20.Text &= "AA cyclone are compact and high efficient" & vbCrLf
        TextBox20.Text &= "AA300 efficiency ~ AC300, capacity 100, ξ= 8.26" & vbCrLf
        TextBox20.Text &= "AA425 efficiency ~ AC435, capacity 70, ξ= 7.24" & vbCrLf
        TextBox20.Text &= "AA600 efficiency ~ AC600, capacity 50, ξ= 8.0" & vbCrLf
        TextBox20.Text &= "AA850 efficiency ~ AC850, capacity 35, ξ= 8.0" & vbCrLf
        TextBox20.Text &= "Capacity step is 0.5*sqrt(2)= 0.71" & vbCrLf & vbCrLf
        TextBox20.Text &= "Load above 5 gr/m3 is considered a high load" & vbCrLf
        TextBox20.Text &= "Cyclones can not choke" & vbCrLf

        TextBox25.Text = "Particulate Matter PM10" & vbCrLf
        TextBox25.Text &= "All particle smaller then 10mu" & vbCrLf
        TextBox25.Text &= "PM10 can travel between 100 meter to 50 km" & vbCrLf

        TextBox27.Text = "Van Tongeren USA" & vbCrLf
        TextBox27.Text &= "The multicyclones have a 168mm diameter instead of 300mm" & vbCrLf
        TextBox27.Text &= " " & vbCrLf

        TextBox46.Text = "The shape of the particle is influencing the separation results" & vbCrLf
        TextBox46.Text &= "The original test are performed with sphere shaped products" & vbCrLf
        TextBox46.Text &= "E.g. plate shaped particles are more difficult to remove due to the" & vbCrLf
        TextBox46.Text &= "different surface to ratio" & vbCrLf

        TextBox47.Text = "Applications" & vbCrLf
        TextBox47.Text &= "Fly catcher Venezuela before a gasturbine" & vbCrLf
        TextBox47.Text &= "Spark catcher (metalic particles)" & vbCrLf
        TextBox47.Text &= "Droplet catcher" & vbCrLf
        TextBox47.Text &= "Patato Starch" & vbCrLf

        TextBox49.Text = "Cyclone feed sampling method" & vbCrLf
        TextBox49.Text &= "Use isokinetic sampling see BS3405, BS6069, ISO9096" & vbCrLf
        TextBox49.Text &= "Note a sample is always better then NO sample" & vbCrLf

        TextBox50.Text = "Notes" & vbCrLf
        TextBox50.Text &= "Group number= There are 11 PSD input groups" & vbCrLf
        TextBox50.Text &= "PSD= Particle Size Distribution (diameter class)" & vbCrLf
        TextBox50.Text &= "PSD diff= difference this and previous PSD dia class" & vbCrLf
        TextBox50.Text &= "Loss overall= Not caught, passed trough cyclone" & vbCrLf
        TextBox50.Text &= "Loss abs= Absolut loss in this diameter class" & vbCrLf
        TextBox50.Text &= "Loss abs corr= Correct. for dust load and discharge pipe" & vbCrLf


        TextBox126.Text = "Calculation method" & vbCrLf
        TextBox126.Text &= "The customer delivers a particle size distribution with minimum 10 fractions " & vbCrLf
        TextBox126.Text &= "Use the Rosin Rammler formula to draw a cumulative particle distribution with the given 10 fractions" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "The cyclone influx is divided into 110 particle fractions based on the "
        TextBox126.Text &= "diameter (the smallest to the largest particle)" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "Start with the smallest diameter fraction of stage 2" & vbCrLf
        TextBox126.Text &= "Each fraction has a loss calculation based on the Weibull formula" & vbCrLf
        TextBox126.Text &= "Determine Stokes for the mean particle and find m, k, a for the cyclone used." & vbCrLf
        TextBox126.Text &= "Note m, k, a are different for particles larger and smaller than the critical particle diameter" & vbCrLf
        TextBox126.Text &= "The critical diameter is given by the cyclone model (AC300, AC350 etc)" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "Determine for each of the 110 particle fractions diameter -minimum [mu], -maximum [mu]," & vbCrLf
        TextBox126.Text &= " -average [mu], Stokes number and factor m, k, a for the chosen cyclone" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "Determine for each Group (in this case 1..10), d1(smallest diameter of the group), d2 (biggest diam.)," & vbCrLf
        TextBox126.Text &= "p1= cumulative weight group lower limit, p2= cum. weight upper limit, the cumulative weight group range is 0.99990-0.001" & vbCrLf
        TextBox126.Text &= "Determine the loss number 0-1 (zero (0) loss for large particles and total (1.0) loss for tiny particles) " & vbCrLf

        TextBox126.Text &= "Now add corrections for the dust load and and the center discharge tube length (normally we use 1.0)" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "We can now calculate the absolute loss per fraction, loss*psd_diff= loss_abs" & vbCrLf
        TextBox126.Text &= "psd_dif summed over de 110 fractions should approach the 1.0 (100%)." & vbCrLf
        TextBox126.Text &= "Absolute Emission= Total-loss-ratio * inlet-dust-load" & vbCrLf
        TextBox126.Text &= "" & vbCrLf
        TextBox126.Text &= "Note 1) particles <0.5-0.7 mu cannot be captured due to physical mechanisms greater than the centrifugal force."

        TextBox145.Text = "Spray dryer outlet pressure range is between 0 and -5 mbar" & vbCrLf
        TextBox145.Text &= "Dairy industry requires 5-10 mu stack exhaust"
        TextBox145.Text &= "Fat accumulation on the cyclone inlet wall (called Caramelization, Per Simonsen)"
        TextBox145.Text &= "Dairy specialists Per Simonsen, Suresh Sundararaj"

        TextBox147.Text = "Cyclone is a excellent Spark arrestor" & vbCrLf
        TextBox147.Text &= "Use in front of filter or silo" & vbCrLf
        TextBox147.Text &= "Also used in suction of gas-turbines to catch flies" & vbCrLf

        TextBox149.Text = "Gluten separation with a Cyclone" & vbCrLf &
        "Test have show that the particles are getting electrically charged and stick to the vessel wall" & vbCrLf

        TextBox152.Text = "Particle Density and Bulk (Volumetric) Density" & vbCrLf &
        "Rule of thumb for foodstuffs, Particle Density= 2 x Bulk density" & vbCrLf &
        "Use: Starch 1500 kg/m2, Protein 1200 kg/m3" & vbCrLf

        TextBox153.Text = "Spray dried product are fragile and may breakup See project P10.1070" & vbCrLf &
        "Inlet speed cyclone is then limited to 16 m/s " & vbCrLf &
        "Strong product inlet speed is 25 m/s" & vbCrLf &
        "Maltodextrine  " & vbCrLf

        TextBox174.Text = "AVEBE" & vbCrLf &
        "Many type of Potato starch" & vbCrLf &
        "Use cyclone as 1 stage before AA850 to prevent blocking"

        TextBox187.Text = "Log" & vbCrLf &
        "15-02-2022, Bugfix printout gram->mg/Nm3" & vbCrLf &
        "10-02-2022, Lognormal Distribution tool added" & vbCrLf &
        "13-11-2021, Now .NET framework 4.8" & vbCrLf &
        "09-11-2021, Tab Steel, Stress# 1 and stress #2 added still under construction" & vbCrLf &
        "02-11-2021, Checkbox added @ viscosity to enable manual input" & vbCrLf &
        "29-10-2021, Bugfix cyclone outside area and weight" & vbCrLf &
        "30-09-2021, Emissions now in milli-gram instead of gram" & vbCrLf &
        "30-09-2021, Inlet flow in Nm3/h added" & vbCrLf &
        "29-09-2021, Mol weight -> density inlet gas added" & vbCrLf &
        "11-05-2021, Bugfix Correction-factor dia. discharge pipe" & vbCrLf &
        "11-05-2021, General code cleanup" & vbCrLf &
        "07-05-2021, PSD Corn starch cleaned up" & vbCrLf &
        "07-05-2021, Chart added Volume percentage vs Diameter" & vbCrLf &
        "07-05-2021, Emission in [gr/Nm3] added to report" & vbCrLf &
        "07-05-2021, Bugfix retrieve project" & vbCrLf &
        "06-05-2021, Chickpea starch added Source Denmark" & vbCrLf &
        "07-04-2021, Commercial data elaborated" & vbCrLf &
        "01-04-2021, PSD Corm added, as difficult product" & vbCrLf &
        "16-03-2021, Bug fix clear the grid" & vbCrLf &
        "07-02-2021, Bug fix input" & vbCrLf &
        "09-02-2021, Invert input button added" & vbCrLf &
        "09-02-2021, Emission in gr/Nm3 added" & vbCrLf &
        "10-02-2021, Emmision on kg/h added " & vbCrLf &
        "10-02-2021, Potato PSD added" & vbCrLf &
        "25-02-2021, AKO rvs kg price 11.00 E/kg "

        TextBox227.Text =
        "Important note" & vbCrLf &
        "The yield strength follows EN 10028-2:2009 (mild steel)" & vbCrLf &
        "and EN 10028-7:2016 for stainless steel at given temperatureerature." & vbCrLf &
        "Safety factors follow the Eurocode" & vbCrLf & vbCrLf &
        "EN 14460:2006, Explosion resistand design follow EN 13445" & vbCrLf &
        "for Explosion-Pressure-Shock-Resistant design stress multiplied bu 1.5"

        Me.Size = New System.Drawing.Size(1305, 906)

        Build_clear_dgv6()              'PSD input grid
        Calc_sequence()
        Design_stress()

        init = True                     'init is now done
        Debug.WriteLine("789  " & _input.GetLength(0).ToString)
        Draw_chart5_weibull()
        Calc_plot_Distribution()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles button1.Click, TabPage1.Enter, numericUpDown2.ValueChanged, NumericUpDown1.ValueChanged, numericUpDown5.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown18.ValueChanged, ComboBox1.SelectedIndexChanged, NumericUpDown4.ValueChanged, NumericUpDown34.ValueChanged, NumericUpDown33.ValueChanged, ComboBox2.SelectedIndexChanged, NumericUpDown43.ValueChanged, NumericUpDown22.ValueChanged, CheckBox3.CheckedChanged, CheckBox2.CheckedChanged, NumericUpDown3.ValueChanged, CheckBox20.CheckedChanged, NumericUpDown6.ValueChanged
        Calc_sequence()
    End Sub

    Private Sub Get_input_calc_1(ks As Integer)
        Dim db1 As Double           'Body diameter stage #1
        Dim db2 As Double           'Body diameter stage #2
        Dim words As String()

        '===Input parameter ===
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
        Dim lmp1, lmp2 As Double        'Total length

        '====  Inlet Conditions stage #1 ====
        _cees(ks).Temp = NumericUpDown18.Value                         'Temperature [c] 
        _cees(ks).p1_abs = NumericUpDown19.Value * 100.0 + 101325      'Pressure [Pa]

        Dim R_gas As Double = 1000 * _Gasconstant / NumericUpDown3.Value                '[J.kg.K]  
        _cees(ks).Ro_gas1_Am3 = _cees(ks).p1_abs / (R_gas * (_cees(ks).Temp + 273.15))  '[kg/Am3]
        _cees(ks).Flow_air_kgh = NumericUpDown1.Value * _cees(ks).Ro_gas1_Am3           '[kg/h] air

        '------ dust load NORMAL conditions ----
        _cees(ks).Ro_gas1_Nm3 = Calc_Normal_density(_cees(ks).Ro_gas1_Am3, _cees(ks).p1_abs, _cees(ks).Temp)

        '--------- ratio  Nm2 and Am3 ---------
        ratio = _cees(ks).Ro_gas1_Nm3 / _cees(ks).Ro_gas1_Am3
        _cees(ks).dust1_Am3 = NumericUpDown4.Value                                      'dust gram/Am3
        _cees(ks).dust1_Nm3 = _cees(ks).dust1_Am3 * ratio                               'dust gram/Nm3

        '--------- present -------------
        TextBox190.Text = _cees(ks).Ro_gas1_Am3.ToString("F3")                          '[kg/Am3]
        TextBox132.Text = _cees(ks).dust1_Nm3.ToString("F3")                            'gram/Nm3
        TextBox129.Text = _cees(ks).Ro_gas1_Nm3.ToString("F3")                          'kg/Nm3 inlet gas
        TextBox191.Text = (NumericUpDown1.Value / ratio).ToString("F0")                 'Nm3/h inlet gas
        TextBox136.Text = _cees(ks).Flow_air_kgh.ToString("F0")                         '[kg/h] air


        If (ComboBox1.SelectedIndex > -1) AndAlso (ComboBox2.SelectedIndex > -1) Then 'Prevent exceptions
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

            _cees(ks).Noc1 = NumericUpDown20.Value 'Parallel cyclones #1
            _cees(ks).Noc2 = NumericUpDown33.Value 'Parallel cyclones #2

            db1 = numericUpDown5.Value                  '[m] Body diameter stage #1
            db2 = NumericUpDown34.Value               '[m] Body diameter stage #2
            _cees(ks).db1 = db1                             '[m] Body diameter
            _cees(ks).db2 = db2                             '[m] Body diameter
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

            ro_solid = numericUpDown2.Value                   '[kg/m3]

            If CheckBox20.Checked Then  'Calc temperature based
                visco = Air_visco(NumericUpDown18.Value)        '[cPoise]
                NumericUpDown6.Enabled = False
                NumericUpDown6.BackColor = Color.White
                NumericUpDown6.Value = CDec(visco)
            Else
                visco = NumericUpDown6.Value
                NumericUpDown6.Enabled = True
                NumericUpDown6.BackColor = Color.Yellow
            End If


            '=========== Stage #1 ==============
            _cees(ks).inv1 = _cees(ks).Flow1 / (_cees(ks).inb1 * _cees(ks).inh1)    '[m/s]
            _cees(ks).outv1 = _cees(ks).Flow1 / ((PI / 4) * _cees(ks).dout1 ^ 2)    '[m/s]

            '=============== Silly values =======
            If _cees(ks).inv1 < 1 Then _cees(ks).inv1 = 1               '[m/s] Prevent silly results
            If _cees(ks).inv1 > 60 Then _cees(ks).inv1 = 60             '[m/s] Prevent silly results

            If ComboBox1.Items.Count > 0 Then    'Cyclone selection
                words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
                wc_air1 = CDbl(words(8))        'Resistance Coefficient air
                wc_dust1 = CDbl(words(9))       'Resistance Coefficient dust
            Else
                MessageBox.Show("Error 763")
            End If
            _cees(ks).dpgas1 = 0.5 * _cees(ks).Ro_gas1_Am3 * _cees(ks).inv1 ^ 2 * wc_air1      '[Pa]
            _cees(ks).dpdust1 = 0.5 * _cees(ks).Ro_gas1_Am3 * _cees(ks).inv1 ^ 2 * wc_dust1    '[Pa]
            _cees(ks).p2_abs = _cees(ks).p1_abs - _cees(ks).dpgas1                              '[P_abs (inlet stage #2)]


            '=========== Inlet Stage #2 ==============
            _cees(ks).Ro_gas2_Am3 = _cees(ks).Ro_gas1_Am3 * _cees(ks).p2_abs / _cees(ks).p1_abs  '[kg/Am3] Inlet stage #2 
            _cees(ks).Ro_gas2_Nm3 = Calc_Normal_density(_cees(ks).Ro_gas2_Am3, _cees(ks).p2_abs, _cees(ks).Temp)

            _cees(ks).Flow2 = _cees(ks).FlowT / (3600 * _cees(ks).Noc2)                 '[Am3/s] Air flow per cyclone 

            '------ compensate the flow for the pressure loss over stage 1 ---
            _cees(ks).Flow2 *= _cees(ks).p1_abs / _cees(ks).p2_abs                      '[Am3/s] Air flow per cyclone 

            '---- Compensate for the Speed for the pressure loss in stage #1 ----
            _cees(ks).inv2 = _cees(ks).Flow2 / (_cees(ks).inb2 * _cees(ks).inh2)         '[m/s]
            _cees(ks).outv2 = _cees(ks).Flow2 / ((PI / 4) * _cees(ks).dout2 ^ 2)        '[m/s]

            '----------- Pressure loss cyclone stage #2----------------------
            words = rekenlijnen(ComboBox2.SelectedIndex).Split(CType(";", Char()))
            wc_air2 = CDbl(words(8))
            wc_dust2 = CDbl(words(9))

            _cees(ks).dpgas2 = 0.5 * _cees(ks).Ro_gas2_Am3 * _cees(ks).inv2 ^ 2 * wc_air2
            _cees(ks).dpdust2 = 0.5 * _cees(ks).Ro_gas2_Am3 * _cees(ks).inv2 ^ 2 * wc_dust2

            '----------- Cyclone stage #2 Outlet conditions ----------------------
            _cees(ks).p3_abs = _cees(ks).p2_abs - _cees(ks).dpgas2
            _cees(ks).Ro_gas3_Am3 = _cees(ks).Ro_gas1_Am3 * _cees(ks).p3_abs / _cees(ks).p1_abs  '[kg/Am3]
            _cees(ks).Ro_gas3_Nm3 = Calc_Normal_density(_cees(ks).Ro_gas3_Am3, _cees(ks).p3_abs, _cees(ks).Temp)


            '----------- 1/2 cone apex #1-----------
            kortezijde = (db1 - _cyl1_dim(12) * db1) * 0.5          '[m]
            langezijde = _cyl1_dim(11) * db1                        '[m]
            halfconeapex1 = Atan(kortezijde / langezijde)           '[rad]
            halfconeapex1 = halfconeapex1 / (PI / 2) * 90.0         '[degree]

            '----------- 1/2 cone apex #2-----------
            kortezijde = (db2 - _cyl1_dim(12) * db2) * 0.5          '[m]
            langezijde = _cyl1_dim(11) * db2                        '[m]
            halfconeapex2 = Atan(kortezijde / langezijde)           '[rad] 
            halfconeapex2 = halfconeapex2 / (PI / 2) * 90.0         '[degree]

            '----------- stof belasting ------------
            kgs = _cees(ks).Flow1 * _cees(ks).dust1_Am3 / 1000      '[kg/s/cycloon]
            kgh = kgs * 3600.0                                      '[kg/h/cycloon]
            _cees(ks).dust1_in_kgh = kgh * _cees(ks).Noc1           '[kg/h] dust inlet

            '----------- K_stokes-----------------------------------
            _cees(ks).Kstokes1 = Sqrt(db1 * 2000 * visco * 16 / (ro_solid * 0.0181 * _cees(ks).inv1))
            _cees(ks).Kstokes2 = Sqrt(db2 * 2000 * visco * 16 / (ro_solid * 0.0181 * _cees(ks).inv2))

            '----------- presenteren ----------------------------------
            TextBox128.Text = _cees(ks).Ro_gas2_Nm3.ToString("F3")      '[kg/Nm3] density
            TextBox183.Text = _cees(ks).Ro_gas3_Nm3.ToString("F3")      '[kg/Nm3] density

            TextBox36.Text = (_cees(ks).FlowT / 3600).ToString("F2")    '[m3/s] flow
            TextBox177.Text = _cees(ks).dust1_in_kgh.ToString("F1")     '[kg/h] dust


            If (ComboBox1.SelectedIndex = 9) Then
                groupBox3.Visible = False
            Else
                groupBox3.Visible = True
            End If

            lmp1 = ((_cyl1_dim(10) + _cyl1_dim(11) + 3.0 * _cyl1_dim(12)) * db1)  'Height Cyclone stage 1
            lmp2 = ((_cyl2_dim(10) + _cyl2_dim(11) + 3.0 * _cyl2_dim(12)) * db2)  'Height Cyclone stage 2

            '----------- presenteren afmetingen AC cyclonen in [m] ----------------------------
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
            TextBox150.Text = halfconeapex1.ToString("F1")           '1/2 cone apex
            TextBox44.Text = lmp1.ToString("F3")                     'L+M+3P


            TextBox84.Text = (_cees(ks).inh2).ToString("F3")         'inlaat breedte
            TextBox85.Text = (_cees(ks).inb2).ToString("F3")         'Inlaat hoogte
            TextBox86.Text = (_cyl2_dim(3) * db2).ToString("F3")     'Inlaat lengte
            TextBox87.Text = (_cyl2_dim(4) * db2).ToString("F3")     'Inlaat hartmaat
            TextBox88.Text = (_cyl2_dim(5) * db2).ToString("F3")     'Inlaat afschuining
            TextBox89.Text = (_cyl2_dim(6) * db2).ToString("F3")     'Uitlaat keeldia inw.
            TextBox90.Text = (_cyl2_dim(7) * db2).ToString("F3")     'Uitlaat flensdiameter inw.
            TextBox91.Text = (_cyl2_dim(8) * db2).ToString("F3")     'Lengte insteekpijp inw.
            TextBox92.Text = (_cyl2_dim(9) * db2).ToString("F3")     'Lengte romp + conus
            TextBox93.Text = (_cyl2_dim(10) * db2).ToString("F3")    'Lengte romp
            TextBox94.Text = (_cyl2_dim(11) * db2).ToString("F3")    'Lengte çonus
            TextBox95.Text = (_cyl2_dim(12) * db2).ToString("F3")    'Dia_conus / 3P-pijp
            TextBox96.Text = (_cyl2_dim(13) * db2).ToString("F3")    'Lengte 3P-pijp
            TextBox97.Text = (_cyl2_dim(14) * db2).ToString("F3")    'Lengte 3P conus
            TextBox98.Text = (_cyl2_dim(15) * db2).ToString("F3")    'Kleine dia 3P-conus
            TextBox151.Text = halfconeapex2.ToString("F1")           '1/2 cone apex
            TextBox45.Text = lmp2.ToString("F3")                     'L+M+3P

            TextBox113.Text = (_cees(ks).Flow1 * 3600.0).ToString("0")    '[Am3/s] Cycloone Flow
            TextBox112.Text = (_cees(ks).Flow2 * 3600.0).ToString("0")    '[Am3/s] Cycloone Flow

            TextBox16.Text = _cees(ks).inv1.ToString("0.0")             'inlaat snelheid
            TextBox80.Text = _cees(ks).inv2.ToString("0.0")             'inlaat snelheid

            TextBox17.Text = _cees(ks).dpgas1.ToString("0")             '[Pa] Pressure loss inlet-gas
            TextBox79.Text = _cees(ks).dpgas2.ToString("0")             '[Pa] Pressure loss inlet-gas
            TextBox189.Text = ((_cees(ks).dpgas1 + _cees(ks).dpgas2) / 100).ToString("F1") '[mbar]

            TextBox48.Text = _cees(ks).dpdust1.ToString("0")            '[Pa] Pressure loss inlet-dust
            TextBox76.Text = _cees(ks).dpdust2.ToString("0")            '[Pa] Pressure loss inlet-dust

            TextBox22.Text = _cees(ks).outv1.ToString("0.0")            'uitlaat snelheid
            TextBox77.Text = _cees(ks).outv2.ToString("0.0")            'uitlaat snelheid

            TextBox23.Text = _cees(ks).Kstokes1.ToString("F4")          'Stokes waarde stage#1
            TextBox78.Text = _cees(ks).Kstokes2.ToString("F4")          'Stokes waarde stage#2

            TextBox37.Text = _cees(ks).db1.ToString                     '[m] Cycloone diameter stage#1
            TextBox74.Text = _cees(ks).db2.ToString                     '[m] Cycloone diameter stage#2

            TextBox38.Text = CType(ComboBox1.SelectedItem, String)      'Cycloon type
            TextBox73.Text = CType(ComboBox2.SelectedItem, String)      'Cycloon type

            '---------- Pressure abs --------------
            TextBox131.Text = _cees(ks).p1_abs.ToString("F0")           '[Pa abs], inlet stage 1
            TextBox130.Text = _cees(ks).p2_abs.ToString("F0")           '[Pa abs], inlet stage 2
            TextBox181.Text = _cees(ks).p3_abs.ToString("F0")           '[Pa abs], outlet stage 2

            '---------- Density --------------
            TextBox75.Text = _cees(ks).Ro_gas1_Am3.ToString("F3")       '[kg/Am3]
            TextBox19.Text = _cees(ks).Ro_gas2_Am3.ToString("F3")       '[kg/Am3]
            TextBox182.Text = _cees(ks).Ro_gas3_Am3.ToString("F3")      '[kg/Am3]

            '---------- Check inlet speed [m/s] stage #1---------------
            TextBox16.BackColor = If(_cees(ks).inv1 < 12 OrElse _cees(ks).inv1 > 32, Color.Red, Color.LightGreen)

            '---------- Check inlet speed stage #2---------------
            TextBox80.BackColor = If(_cees(ks).inv2 < 12 OrElse _cees(ks).inv2 > 32, Color.Red, Color.LightGreen)

            '---------- Check dp [pa] stage #1---------------
            TextBox17.BackColor = If(_cees(ks).dpgas1 > 3000.0, Color.Red, Color.LightGreen)

            '---------- Check dp stage #2---------------
            TextBox79.BackColor = If(_cees(ks).dpgas2 > 3000.0, Color.Red, Color.LightGreen)

            '--------- Get Inlet korrel-groep data ----------
            'Save data of screen into the _cees array
            'Fill_array_from_screen(CInt(NumericUpDown30.Value))

            TextBox39.Text = kgh.ToString("F0")                 'Stof inlet [kg/(h.cyclone)]
            TextBox40.Text = _cees(ks).dust1_in_kgh.ToString("F0")             'Dust inlet [kg/h] 
            TextBox71.Text = _cees(ks).dust1_Am3.ToString("F3") 'Dust inlet [g/Am3]
        End If
    End Sub
    Private Sub Present_Datagridview1(ks As Integer)

        With DataGridView1
            '--------- HeaderText --------------------
            .Columns(0).HeaderText = "Dia class"
            .Columns(1).HeaderText = "Cyc feed psd cum"
            .Columns(2).HeaderText = "Cyc feed psd diff"
            .Columns(3).HeaderText = "Loss [%] of feed"
            .Columns(4).HeaderText = "Loss abs [%]"
            .Columns(5).HeaderText = "Loss psd cum"
            .Columns(6).HeaderText = "Catch abs"
            .Columns(7).HeaderText = "Catch psd cum"
            .Columns(8).HeaderText = "Grade class eff."

            Calc_Datagridview1(ks)
            '=========  sum is required for column 5 calculation ===============
            k41 = 0
            For h = 0 To 22
                k41 += CDbl(.Rows(h).Cells(4).Value)    'tot_catch_abs[%]
            Next
            '===================================================================

            .AutoResizeColumns()
        End With
    End Sub
    Private Sub Calc_Datagridview1(ks As Integer)
        'This is a summary of the real calulation 23 lines long
        '
        '==== stage #1 + stage #2 ====
        Dim h18, h19 As Double
        Dim j18, i18 As Double
        Dim l18, k19 As Double
        Dim k18 As Double
        Dim m18, n17_oud, n18 As Double
        Dim tot_catch_abs As Double
        Dim o18 As Double
        Dim tt As Double


        With DataGridView1
            For h = 0 To 22
                .Rows(h).Cells(0).Value = _cees(ks).stage1(h * 5).d_ave.ToString("F3")         'diameter
                .Rows(h).Cells(1).Value = _cees(ks).stage1(h * 5).psd_cum_pro.ToString("F3")   'feed psd cum

                If h > 0 Then
                    h18 = CDbl(.Rows(h - 1).Cells(1).Value)
                Else
                    h18 = 100.0
                End If

                h19 = CDbl(.Rows(h).Cells(1).Value)   'feed psd cum
                .Rows(h).Cells(2).Value = (h18 - h19).ToString("F3")   'feed psd diff

                '========= (column 3) loss ===============
                If CheckBox2.Checked Then
                    .Rows(h).Cells(3).Value = (_cees(ks).stage1(h * 5).loss_overall_C * 100.0).ToString("F3")
                Else
                    .Rows(h).Cells(3).Value = (_cees(ks).stage1(h * 5).loss_overall * 100.0).ToString("F3")
                End If

                i18 = CDbl(.Rows(h).Cells(2).Value) 'feed psd diff
                j18 = CDbl(.Rows(h).Cells(3).Value) 'loss % Of feed
                .Rows(h).Cells(4).Value = (i18 * j18 / 100.0).ToString("F3") 'Loss abs [%]

                '=========  (column 4) Loss abs [%] ===============
                'If h > 0 Then
                '    l18 = CDbl( .Rows(h - 1).Cells(5).Value)
                'Else
                '    l18 = 100
                'End If
                k19 = CDbl(.Rows(h).Cells(4).Value)   'Loss abs [%]

                '=========  (column 5) Loss psd cum ===============
                If h > 0 Then
                    l18 = CDbl(.Rows(h - 1).Cells(5).Value)
                Else
                    l18 = 100.0
                End If
                tt = (l18 - 100.0 * k19 / k41)
                If tt < 0 Then tt = 0.0           'Prevent negative numbers
                .Rows(h).Cells(5).Value = tt.ToString("F3")

                '============= (column 6) Loss abs [%] ===================
                k18 = CDbl(.Rows(h).Cells(4).Value)   'Loss abs [%]
                m18 = (i18 - k18)
                .Rows(h).Cells(6).Value = m18.ToString("F3") 'Catch abs

                '=============  (column 7) Catch psd cum  ===================
                Double.TryParse(TextBox59.Text, tot_catch_abs)      'tot_catch_abs[%]
                If h > 0 Then
                    n17_oud = CDbl(.Rows(h - 1).Cells(7).Value)
                    n18 = n17_oud - m18 / (tot_catch_abs / 100.0)
                Else
                    n18 = 100.0
                End If
                n18 = CDbl(IIf(n18 < 0, 0, n18))        'prevent silly results
                .Rows(h).Cells(7).Value = n18.ToString("F3") 'Catch psd cum

                '=========  (column 8) Efficiency ===============
                o18 = 100.0 - j18
                .Rows(h).Cells(8).Value = o18.ToString("F3") 'Grade eff.
            Next h
        End With
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

        _cees(c_nr).Quote_no = TextBox28.Text                   'Quote number
        _cees(c_nr).Tag_no = TextBox29.Text                     'The Tag number
        _cees(c_nr).case_name = TextBox53.Text                  'The case name

        _cees(c_nr).FlowT = NumericUpDown1.Value          'Air flow [Am3/hr]
        _cees(c_nr).dust1_Am3 = NumericUpDown4.Value      'Dust inlet [g/Am3] 
        _cees(c_nr).Ct1 = ComboBox1.SelectedIndex         'Cyclone type Stage #1
        _cees(c_nr).Ct2 = ComboBox2.SelectedIndex         'Cyclone type Stage #2
        _cees(c_nr).Noc1 = NumericUpDown20.Value          'Cyclone no in parallel #1
        _cees(c_nr).Noc2 = NumericUpDown33.Value          'Cyclone no in parallel #2
        _cees(c_nr).db1 = numericUpDown5.Value            '[m] Diameter cyclone Stage #1
        _cees(c_nr).db2 = NumericUpDown34.Value           '[m] Diameter cyclone Stage #2
        If TextBox190.Text = "" Then
            TextBox190.Text = "99"
        End If
        _cees(c_nr).ro_gas = CDbl(TextBox190.Text)             'Density [kg/m3]
        _cees(c_nr).ro_solid = numericUpDown2.Value       'Density [kg/m3]
        _cees(c_nr).Temp = NumericUpDown18.Value          'Temperature [c]
        _cees(c_nr).p1_abs = 101325 + (NumericUpDown19.Value * 100)        'Pressure [Pa abs]
    End Sub
    Private Sub Read_dgv6_Calc_Class_load()

        '==== read dgv6 and do not trip on "-" ====
        '[mu] Class upper particle diameter limit diameter
        'Percentale van de inlaat stof belasting [%]

        Dim a, b As Double
        Dim st1, st2 As String

        For row = 0 To DataGridView6.Rows.Count - 1
            If Not IsNothing(DataGridView6.Rows(row).Cells(0).Value) AndAlso Not IsNothing(DataGridView6.Rows(row).Cells(1).Value) Then
                st1 = DataGridView6.Rows(row).Cells(0).Value.ToString
                st2 = DataGridView6.Rows(row).Cells(1).Value.ToString

                Double.TryParse(st1, a)
                Double.TryParse(st2, b)

                If a > 0 AndAlso b > 0 Then
                    _input(row).dia_big = a
                    _input(row).class_load = b / 100.0
                Else
                    _input(row).dia_big = 0
                    _input(row).class_load = 0
                End If
            End If
        Next
    End Sub

    '-------- Bereken het verlies getal NIET gecorrigeerd -----------
    '----- de input is de GEMIDDELDE korrel grootte-----------
    Private Function Calc_verlies(korrel_g As Double, stokes As Double, stage As Integer) As Double
        Dim words As String()
        Dim dia_Kcrit, fac_m, fac_a, fac_k As Double
        Dim verlies As Double
        Dim dia_K As Double

        If (ComboBox1.Items.Count > 0 AndAlso ComboBox2.Items.Count > 0) Then
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
                verlies = 1.0       '100% loss (very small particle)
            End If
        Else
            verlies = 1.0           '100% loss (very small particle)
        End If
        Return (verlies)
    End Function

    '-------- Bereken het verlies getal GECORRIGEERD -----------
    '----- de input is de GEMIDDELDE korrel grootte-----------

    Private Sub Calc_verlies_corrected(ByRef grp As GvG_Calc_struct, stage As Integer)
        Dim cor1, cor2 As Double

        If stage > 2 OrElse stage < 1 Then MessageBox.Show("Problem in Line 1034")  '----- check input ----

        If (ComboBox1.Items.Count > 0) AndAlso ComboBox2.Items.Count > 0 Then
            If (stage = 1) Then                         'Stage #1 cyclone
                cor1 = 1 / CDbl(NumericUpDown22.Value)      'Correctie insteek pijp stage #1
                Double.TryParse(TextBox55.Text, cor2)   'Hoge stof belasting correctie acc VT-UK
            Else                                        'Stage #2 cyclone
                cor1 = 1 / CDbl(NumericUpDown43.Value)      'Correctie insteek pijp stage #2
                Double.TryParse(TextBox67.Text, cor2)   'Hoge stof belasting correctie acc VT-UK
            End If
            grp.loss_overall_C = grp.loss_overall ^ (cor1 * cor2)
        Else
            MsgBox("Cyclone stage 1 ot 2 not selected")
        End If
    End Sub

    'Note dp(95) meaning with this diameter 95% is lost
    'Calculate the diameter at which qq% is lost
    'Separation depends on Stokes
    Private Function Calc_dia_particle(qq As Double, stokes As Double, stage As Integer) As Double
        Dim dia_result As Double
        Dim words As String()
        Dim dia_Kcrit As Double
        Dim d1, d2 As Double
        Dim cor1, cor2 As Double        'Insteek pijp diameter
        Dim fac_m1, fac_k1, fac_a1 As Double
        Dim fac_m2, fac_k2, fac_a2 As Double

        '--- check input ----
        If qq > 1 Then MessageBox.Show("Loss > 100% is impossible, Line 486, qq= " & qq.ToString)
        If stage > 2 OrElse stage < 1 Then MessageBox.Show("Problem in Line 708")

        '----- Insteek pijp corectie correctie -------
        If (stage = 1) Then                         'Cyclone 1# stage
            cor1 = 1 / CDbl(NumericUpDown22.Value)  'Correctie insteek pijp stage #1
            Double.TryParse(TextBox55.Text, cor2)   'Hoge stof belasting correctie acc VT-UK
            words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
        Else                                        'Cyclone 2# stage
            cor1 = 1 / CDbl(NumericUpDown43.Value)  'Correctie insteek pijp stage #2
            Double.TryParse(TextBox67.Text, cor2)   'Hoge stof belasting correctie acc VT-UK
            words = rekenlijnen(ComboBox2.SelectedIndex).Split(CType(";", Char()))
        End If

        '-------------- korrelgrootte factoren ------
        dia_Kcrit = CDbl(words(1))   'Is in fact d/K(crit)

        '---- diameter particle kleiner dan de diameter kritisch
        fac_m1 = CDbl(words(2))
        fac_k1 = CDbl(words(3))
        fac_a1 = CDbl(words(4))
        d1 = fac_k1 * stokes * (-Math.Log(qq ^ (1 / (cor1 * cor2)))) ^ (1 / fac_a1) + fac_m1 * stokes

        '---- diameter particle groter dan de diameter kritisch
        fac_m2 = CDbl(words(5))
        fac_k2 = CDbl(words(6))
        fac_a2 = CDbl(words(7))
        d2 = fac_k2 * stokes * (-Math.Log(qq ^ (1 / (cor1 * cor2)))) ^ (1 / fac_a2) + fac_m2 * stokes

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

        '==== prevent silly results =====
        If f < 1 Then f = 1
        If f > 3 Then f = 3

        Return (f)
    End Function

    '---- According to VTK UK -----
    Private Sub Dust_load_correction(ks As Integer)
        Dim f_used1, f_used2 As Double

        '============ stage 1 cyclone ==========
        f_used1 = Calc_dust_load_correction(_cees(ks).dust1_Am3 / 1000)
        '============ stage 2 cyclone ==========
        f_used2 = Calc_dust_load_correction(_cees(ks).dust2_Am3 / 1000)

        '----------- present ------------
        TextBox55.Text = f_used1.ToString("F3")
        TextBox67.Text = f_used2.ToString("F3")
    End Sub

    Public Sub Draw_chart1(ch As Chart, ks As Integer)
        Dim a, b As Double

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
        ch.ChartAreas("ChartArea0").AxisY.Minimum = 0.0      'Loss
        ch.ChartAreas("ChartArea0").AxisY.Maximum = 100.0    'Loss
        ch.ChartAreas("ChartArea0").AxisY.Interval = 10.0    'Interval

        If CheckBox4.Checked Then
            ch.ChartAreas("ChartArea0").AxisX.MinorGrid.Enabled = True
            ch.ChartAreas("ChartArea0").AxisY.MinorGrid.Enabled = True
        End If

        If CheckBox7.Checked Then
            ch.ChartAreas("ChartArea0").AxisX.MinorTickMark.Enabled = True
            ch.ChartAreas("ChartArea0").AxisY.MinorTickMark.Enabled = True
        End If

        ch.ChartAreas("ChartArea0").AxisX.IsLogarithmic = True

        Select Case True
            Case RadioButton1.Checked
                ch.ChartAreas("ChartArea0").AxisX.Minimum = 0.1     'Particle size
                ch.ChartAreas("ChartArea0").AxisX.Maximum = 1000    'Particle size
            Case RadioButton2.Checked
                ch.ChartAreas("ChartArea0").AxisX.Minimum = 1.0     'Particle size
                ch.ChartAreas("ChartArea0").AxisX.Maximum = 10000   'Particle size
            Case Else
                ch.ChartAreas("ChartArea0").AxisX.Minimum = 0.1     'Particle size
                ch.ChartAreas("ChartArea0").AxisX.Maximum = 10000   'Particle size
        End Select

        '----- now start plotting ------------------------

        '----------------------------- Plot Input stars-----------------------
        If CheckBox15.Checked Then
            For i = 0 To (no_PDS_inputs - 1)            'Number of input data points
                a = _input(i).dia_big                   '[mu] Class upper particle diameter limit diameter
                b = _input(i).class_load * 100.0        'Percentale van de inlaat stof belasting
                ch.Series(0).Points.AddXY(a, b)
                ch.Series(0).Points(i).MarkerStyle = MarkerStyle.Star10
                ch.Series(0).Points(i).MarkerSize = 15
            Next
            ch.Series(0).Points(5).Label = "Input"
        End If

        '------ PSD Input Rosin Rammler stage #1 -------------
        If CheckBox14.Checked Then
            For h = 0 To 110 Step 5   'Fill line chart
                a = _cees(ks).stage1(h).dia
                b = _cees(ks).stage1(h).psd_cum_pro
                ch.Series(1).Points.AddXY(a, b)
            Next h
        End If

        '------ Plot Stage #2 output-------------
        '------ Data from DataGridView2 -------------
        If CheckBox16.Checked Then
            For h = 0 To DataGridView2.Rows.Count - 1                   'Fill line chart
                a = CDbl(DataGridView2.Rows(h).Cells(0).Value)        'Particle size
                b = CDbl(DataGridView2.Rows(h).Cells(5).Value)        'Get data eff1
                ch.Series(4).Points.AddXY(a, b)                         'Plot eff #1
            Next h
        End If

        If CheckBox17.Checked Then
            For h = 0 To DataGridView2.Rows.Count - 1                   'Fill line chart
                a = CDbl(DataGridView2.Rows(h).Cells(0).Value)        'Particle size
                b = CDbl(DataGridView3.Rows(h).Cells(5).Value)        'Get data eff2
                ch.Series(5).Points.AddXY(a, b)                         'Plot eff #2
            Next h
        End If

        If CheckBox18.Checked Then
            For h = 0 To DataGridView2.Rows.Count - 1                   'Fill line chart
                a = CDbl(DataGridView4.Rows(h).Cells(0).Value)        'Particle size
                b = CDbl(DataGridView4.Rows(h).Cells(9).Value)        'get data Eff 1&2 [%]
                ch.Series(6).Points.AddXY(a, b)                         'Plot Eff 1&2 [%]
            Next h
        End If

        '----- labels on chart ------------
        If CheckBox16.Checked Then ch.Series(4).Points(30).Label = "Eff #1"
        If CheckBox17.Checked Then ch.Series(5).Points(45).Label = "Eff #2"

        If CheckBox19.Checked Then
            For h = 0 To DataGridView4.Rows.Count - 1                   'Fill line chart
                a = CDbl(DataGridView4.Rows(h).Cells(0).Value)        'Particle size
                b = CDbl(DataGridView4.Rows(h).Cells(5).Value)        'Get data Loss 1
                ch.Series(2).Points.AddXY(a, b)                         'Plot Loss 1
            Next
        End If

        If CheckBox6.Checked Then
            For h = 0 To DataGridView4.Rows.Count - 1                   'Fill line chart
                a = CDbl(DataGridView4.Rows(h).Cells(0).Value)        'Particle size
                b = CDbl(DataGridView3.Rows(h).Cells(14).Value)       'Get data Loss 2
                ch.Series(3).Points.AddXY(a, b)                         'Plot Loss 2
            Next h
        End If

    End Sub

    Private Sub Draw_chart2(ch As Chart, ks As Integer)
        'Small chart on the first tab
        Dim s_points(50, 2) As Double
        Dim h As Integer
        Dim sdia As Integer

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
        ch.ChartAreas("ChartArea0").AxisY.Maximum = 100.0   'Loss
        ch.ChartAreas("ChartArea0").AxisX.Minimum = 0     'Particle size
        ch.ChartAreas("ChartArea0").AxisX.Maximum = 30    'Particle size

        '----- now calc chart points #1 and #2 --------------------------
        Integer.TryParse(TextBox42.Text, sdia)                  'dp(100) stage #1
        s_points(0, 0) = sdia                                   'Particle diameter [mu]
        s_points(0, 1) = 100.0                                  '100% loss
        s_points(0, 2) = 100.0                                  '100% loss

        For h = 1 To 30                                          'Particle diameter [mu]
            s_points(h, 0) = h                                   'Particle diameter [mu]
            s_points(h, 1) = Calc_verlies(h, _cees(ks).Kstokes1, 1) * 100.0  'stage #1 Loss [%]
            s_points(h, 2) = Calc_verlies(h, _cees(ks).Kstokes2, 2) * 100.0  'stage #2 Loss [%]
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
        s_points(0, 0) = 0.0                                        '[kg/Am3] Dust load 0 
        s_points(0, 1) = 1.0                                        '[-] Correction factor

        For h = 1 To 150                                            '[kg/Am3] Dust load 
            s_points(h, 0) = h                                      '[kg/Am3] Dust load 
            s_points(h, 1) = Calc_dust_load_correction(h / 1000)    '[-] Correction factor 
        Next

        '------ now present-------------
        For h = 0 To 150 - 1   'Fill line chart
            ch.Series(0).Points.AddXY(s_points(h, 0), s_points(h, 1))   '
        Next h

        ch.Series(0).Points(6).Label = "Load correction"
    End Sub

    Private Sub Draw_chart4(ch As Chart)
        'Lust Load correction formula
        Dim r_points(no_PDS_inputs, 2) As Double
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

        ch.Titles.Add("PSD selected product")
        ch.ChartAreas("ChartArea0").AxisX.Title = "Particle diameter [mu]"
        ch.ChartAreas("ChartArea0").AxisY.Title = "Volume fraction [%]"
        ch.ChartAreas("ChartArea0").AxisX.Minimum = 0                '
        ' ch.ChartAreas("ChartArea0").AxisX.IsLogarithmic = True

        '----- Calc the gauss shape -----------------------
        r_points(0, 0) = _input(h).dia_big                                  '[mu] 
        r_points(0, 1) = 1.0 - _input(h).class_load                         '[%]  


        For h = 1 To no_PDS_inputs - 1                                          '
            r_points(h, 0) = _input(h).dia_big                                  '[mu] 
            r_points(h, 1) = _input(h - 1).class_load - _input(h).class_load    '[%]

            '===== Diameter is nul then percentage is also nul
            If r_points(h, 0) = 0 Then r_points(h, 1) = 0.0
        Next

        TextBox24.Text = "Present selected PSD" & vbCrLf

        '------ now present-------------
        'https://docs.microsoft.com/en-us/dotnet/standard/base-types/stringbuilder
        Dim sb As New StringBuilder()
        For h = 0 To no_PDS_inputs - 1    'Fill line chart
            sb.AppendFormat("dia={0:F2} , perc = {1:F5}", r_points(h, 0), r_points(h, 1) & vbCrLf)
            ch.Series(0).Points.AddXY(r_points(h, 0), r_points(h, 1))   '
        Next h
        TextBox24.Text &= sb.ToString
        MsgBox(sb.ToString)
    End Sub

    Private Sub Draw_chart5_weibull()
        With Chart5
            .Series.Clear()
            .ChartAreas.Clear()
            .Titles.Clear()
            .ChartAreas.Add("ChartArea0")

            .Series.Add("Series50")
            .Series(0).ChartArea = "ChartArea0"
            .Series(0).ChartType = DataVisualization.Charting.SeriesChartType.Line
            .Series(0).BorderWidth = 2
            .Series(0).IsVisibleInLegend = False

            .Series.Add("Series51")
            .Series(1).ChartArea = "ChartArea0"
            .Series(1).ChartType = DataVisualization.Charting.SeriesChartType.Line
            .Series(1).BorderWidth = 2
            .Series(1).IsVisibleInLegend = False

            .Series.Add("Series52")
            .Series(2).ChartArea = "ChartArea0"
            .Series(2).ChartType = DataVisualization.Charting.SeriesChartType.Line
            .Series(2).BorderWidth = 3
            .Series(2).IsVisibleInLegend = False
            .Series(2).BorderDashStyle = ChartDashStyle.DashDot

            .Titles.Add("Weibull Distribution")
            .ChartAreas("ChartArea0").AxisX.Title = "Particle diameter [mu]"
            .ChartAreas("ChartArea0").AxisX.IsLogarithmic = False
            .ChartAreas("ChartArea0").AxisY.Title = "Exit loss [%]"
            .ChartAreas("ChartArea0").AxisY.Minimum = 0D     'Loss
            .ChartAreas("ChartArea0").AxisY.Maximum = 100D   '[%] weight
            .ChartAreas("ChartArea0").AxisX.Minimum = 0.1D   '[mu] Particle size
            .ChartAreas("ChartArea0").AxisX.Maximum = 100D   '[mu] Particle size
            .ChartAreas("ChartArea0").AxisX.LabelStyle.Format = "F0"

            '------ now present-------------
            For h = 0 To norm_log_dist_pdf.GetLength(0) - 1   'Fill line chart
                If norm_log_dist_pdf(h, 0) > 0 AndAlso norm_log_dist_pdf(h, 1) > 0 Then
                    .Series(0).Points.AddXY(norm_log_dist_pdf(h, 0), norm_log_dist_pdf(h, 1))   'Log normal
                End If
                If norm_log_dist_cdf(h, 0) > 0 AndAlso norm_log_dist_cdf(h, 1) > 0 Then
                    .Series(1).Points.AddXY(norm_log_dist_cdf(h, 0), norm_log_dist_cdf(h, 1))   'Log normal
                End If
            Next h

            '------- PSD input data -----
            Dim a, b As Double
            If CheckBox21.Checked Then
                For h = 0 To DataGridView6.Rows.Count - 1                 'Fill line chart
                    a = CDbl(DataGridView6.Rows(h).Cells(0).Value)        'Particle upper diameter
                    b = CDbl(DataGridView6.Rows(h).Cells(1).Value)        'Cum PSD weight
                    .Series(2).Points.AddXY(a, b)                         'PSD 
                Next
            End If
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles TabPage9.Enter, CheckBox6.CheckedChanged, CheckBox7.CheckedChanged, CheckBox4.CheckedChanged, CheckBox9.CheckedChanged, CheckBox8.CheckedChanged, CheckBox10.CheckedChanged, CheckBox5.CheckedChanged, CheckBox12.CheckedChanged, CheckBox11.CheckedChanged, CheckBox1.CheckedChanged, CheckBox13.CheckedChanged, CheckBox14.CheckedChanged, CheckBox15.CheckedChanged, CheckBox16.CheckedChanged, CheckBox17.CheckedChanged, CheckBox18.CheckedChanged, CheckBox19.CheckedChanged, RadioButton2.CheckedChanged, RadioButton1.CheckedChanged
        Draw_chart2(Chart2, 0)      'Present the results loss curve
    End Sub
    Private Sub Calc_sequence()
        Dim case_nr As Integer = CInt(NumericUpDown30.Value)

        If Update_screen_fram_array_done Then

            Read_dgv6_Calc_Class_load()
            Check_DGV6()

            If ComboBox1.Items.Count > 0 AndAlso ComboBox2.Items.Count > 0 AndAlso init Then

                ProgressBar1.Visible = True
                ProgressBar1.Value = 50
                SuspendLayout()                             'Speedup the program

                '============== ALL calculation is done is Case(0) ========
                '============== other cases are for storage ===============
                _cees(0) = _cees(case_nr)                   'Transfer data to case(0)
                '==========================================================

                If String.Equals(ComboBox1.SelectedItem.ToString, "AA850") Then
                    numericUpDown5.Value = CDec(0.3)        '[m] Diameter
                End If

                If String.Equals(ComboBox2.SelectedItem.ToString, "AA850") Then
                    NumericUpDown34.Value = CDec(0.3)       '[m] Diameter
                End If

                Fill_array_from_screen(0)   'Read input data from sceen
                Dust_load_correction(0)

                Get_input_calc_1(0)         'This is the CASE number
                Calc_part_dia_loss(0)

                For i = 0 To 3
                    Calc_stage1(0)          'Calc according stage #1
                    Calc_stage2(0)          'Calc according stage #2
                Next

                Calc_stage1_2_comb(0)       'Calc stage #1 and stage #2 combined

                Present_loss_grid1(0)       'Present the results stage #1
                Present_loss_grid2(0)       'Present the results stage #2
                Present_Datagridview1(0)    'Present the results stage #1

                Draw_chart1(Chart1, 0)      'Present the results 
                Draw_chart2(Chart2, 0)      'Present the results loss curve
                Screen_contrast()           'White text on ted background 
                ResumeLayout()              'Calcu is done update screen
                Calc_plot_Distribution()
                ProgressBar1.Visible = False
            End If
        End If
        'Debug.WriteLine("Calc_sequence()")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Save the project data to file
        'Return to case 0, save data to file, return to previous selected case
        Save_present_case_to_array()            'Store data in array
        Save_to_disk()
    End Sub
    Private Sub Save_to_disk()
        Dim filename, user As String
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        If TextBox28.Text.Trim.Length = 0 OrElse TextBox29.Text.Trim.Length = 0 Then
            MessageBox.Show("Complete Quote And Tag number")
        Else
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
                MessageBox.Show("Can Not create directory On the VTK intranet (L6286) " & vbCrLf & vbCrLf & ex.Message)
            End Try

            Try
                If Directory.Exists(dirpath_Eng) Then
                    filename = dirpath_Eng & filename 'used at VTK with intranet
                Else
                    filename = dirpath_tmp & filename 'used at VTK with intranet'used at home
                    MessageBox.Show("VTK intranet Not acceasable, saved On " & filename)
                End If

                '--- Delete existing file -------
                If System.IO.File.Exists(filename) Then
                    System.IO.File.Delete(filename)
                End If

                '--- and save new file to disk -------
                Dim fStream As New FileStream(filename, FileMode.CreateNew)
                bf.Serialize(fStream, _cees) ' write to file
                bf.Serialize(fStream, _input) ' write to file
                fStream.Close()
            Catch ex As Exception
                MessageBox.Show("Line 6298, " & ex.Message)  ' Show the exception's message.
            End Try
        End If
    End Sub
    Private Sub Retrieve_from_disk()
        Dim bf As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        ProgressBar1.Visible = True
        OpenFileDialog1.FileName = "Cyclone_select_*.vtk2"

        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_tmp  'used at home
        End If

        OpenFileDialog1.Title = "Open a VTK2 File"
        OpenFileDialog1.Filter = "VTK2 Files|*.vtk2"
        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ProgressBar1.Value = 30
            Dim fStream As New FileStream(OpenFileDialog1.FileName, FileMode.Open) With {
                    .Position = 0 ' reset stream pointer
                    }
            ProgressBar1.Value = 60
            _cees = CType(bf.Deserialize(fStream), Input_struct()) ' read from file
            ProgressBar1.Value = 90
            _input = CType(bf.Deserialize(fStream), Psd_input_struct()) ' read from file
            fStream.Close()

            '==== the length of _input is CHANGED bij Deserialize, return to correct length
            ReDim Preserve _input(no_PDS_inputs + 1)

            Fill_DGV6_from_input_array()
        Else
            TextBox24.Text &= "Retrieved project from disk failed" & vbCrLf
        End If
        ProgressBar1.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        'Retrieve project from disk and goto case nr 0
        Retrieve_from_disk()            'Read from disk to array

        '======= program chokes on this section ===========
        Update_screen_fram_array_done = False
        Update_Screen_from_array(1)     'Refresh screen data case 1
        Update_screen_fram_array_done = True
        '======= done ===========
        Calc_sequence()
    End Sub

    Public Function HardDisc_Id() As String
        'Add system.management as reference !!
        'imports system.management
        Dim tmpStr2 As String = ""
        Dim myScop As New ManagementScope("\\" & Environment.MachineName & "\root\cimv2")
        Dim oQuer As New SelectQuery("Select * FROM WIN32_DiskDrive")

        Dim oResult As New ManagementObjectSearcher(myScop, oQuer)
        Dim oIte As ManagementObject
        Dim oPropert As PropertyData
        For Each oIte In oResult.Get()
            For Each oPropert In oIte.Properties
                If oPropert.Value IsNot Nothing AndAlso oPropert.Name = "SerialNumber" Then
                    tmpStr2 = oPropert.Value.ToString
                    '  Exit For
                End If
            Next
            ' Exit For
        Next
        Return (Trim(tmpStr2))         'Harddisk identification
    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox28.Text.Trim.Length > 0 Then
            Calc_sequence()
            Write_to_word_com()     'Commercial data to Word
        Else
            MessageBox.Show("Enter Quote nummer And Tag, Then Export sizing data To Word")
        End If
    End Sub

    'Write COMMERCIAL data to Word 
    'see https://msdn.microsoft.com/en-us/library/office/aa192495(v=office.11).aspx
    Private Sub Write_to_word_com()
        '  Dim bmp_tab_page1 As New Bitmap(TabPage1.Width, TabPage1.Height)
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara4 As Word.Paragraph

        Dim chart_size As Integer = 65  '% of original picture size
        Dim file_name As String
        Dim row As Integer

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
        oPara1.Range.Text = "VTK Sales, Cyclone sizing"
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
        oTable.Rows(1).Range.Font.Bold = CInt(True)
        row = 2
        oTable.Cell(row, 1).Range.Text = "Project number"
        oTable.Cell(row, 2).Range.Text = TextBox28.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Tag nummer "
        oTable.Cell(row, 2).Range.Text = TextBox29.Text
        row += 1
        oTable.Cell(row, 1).Range.Text = "Print Date"
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
        oTable.Cell(row, 2).Range.Text = TextBox190.Text
        oTable.Cell(row, 3).Range.Text = "[kg/m3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Air viscosity"
        oTable.Cell(row, 2).Range.Text = NumericUpDown6.Value.ToString("F0")
        oTable.Cell(row, 3).Range.Text = "[centi Poise]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Dust load total"
        oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value.ToString("F2")
        oTable.Cell(row, 3).Range.Text = "[gr/Am3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Dust load total"
        oTable.Cell(row, 2).Range.Text = TextBox177.Text
        oTable.Cell(row, 3).Range.Text = "[kg/h]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Dust load stage 1"
        oTable.Cell(row, 2).Range.Text = TextBox39.Text
        oTable.Cell(row, 3).Range.Text = "[kg/h/Cycl]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Dust load stage 2)"
        oTable.Cell(row, 2).Range.Text = TextBox100.Text
        oTable.Cell(row, 3).Range.Text = "[kg/h/Cycl]"

        oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns 
        oTable.Columns(2).Width = oWord.InchesToPoints(1)
        oTable.Columns(3).Width = oWord.InchesToPoints(2)

        oTable.Rows(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '---------------cyclone data-------------------------------
        'Insert a table, fill it with data and change the column widths.
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
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
        oTable.Cell(row, 3).Range.Text = "[m]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "No parallel #1"
        oTable.Cell(row, 2).Range.Text = NumericUpDown20.Value.ToString
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 2
        '----------- stage #2 ---------------
        oTable.Cell(row, 1).Range.Text = "Cyclone type stage #2 "
        oTable.Cell(row, 2).Range.Text = ComboBox2.SelectedItem.ToString
        oTable.Cell(row, 3).Range.Text = "[-]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Body diameter #2"
        oTable.Cell(row, 2).Range.Text = NumericUpDown34.Value.ToString
        oTable.Cell(row, 3).Range.Text = "[m]"
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
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 25, 3)
        oTable.Range.ParagraphFormat.SpaceAfter = 1
        oTable.Range.Font.Size = 10
        oTable.Range.Font.Bold = CInt(False)
        oTable.Rows(1).Range.Font.Bold = CInt(True)
        row = 1
        oTable.Cell(row, 1).Range.Text = "Process data"
        row += 1
        '----------- stage #1 ---------------
        oTable.Cell(row, 1).Range.Text = "Inlet speed stage1 "
        oTable.Cell(row, 2).Range.Text = TextBox16.Text
        oTable.Cell(row, 3).Range.Text = "[m/s]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Outlet speed stage1"
        oTable.Cell(row, 2).Range.Text = TextBox22.Text
        oTable.Cell(row, 3).Range.Text = "[m/s]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Pressure loss stage1"
        oTable.Cell(row, 2).Range.Text = TextBox17.Text
        oTable.Cell(row, 3).Range.Text = "[Pa]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Efficiency stage1"
        oTable.Cell(row, 2).Range.Text = TextBox21.Text
        oTable.Cell(row, 3).Range.Text = "[%]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage1"
        oTable.Cell(row, 2).Range.Text = TextBox18.Text
        oTable.Cell(row, 3).Range.Text = "[mg/Am3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage1"
        oTable.Cell(row, 2).Range.Text = TextBox180.Text
        oTable.Cell(row, 3).Range.Text = "[mg/Nm3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage1"
        oTable.Cell(row, 2).Range.Text = TextBox186.Text
        oTable.Cell(row, 3).Range.Text = "[kg/h]"
        row += 2
        '----------- stage #2 ---------------
        oTable.Cell(row, 1).Range.Text = "Inlet speed stage2 "
        oTable.Cell(row, 2).Range.Text = TextBox80.Text
        oTable.Cell(row, 3).Range.Text = "[m/s]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Outlet speed stage2"
        oTable.Cell(row, 2).Range.Text = TextBox77.Text
        oTable.Cell(row, 3).Range.Text = "[m/s]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Pressure loss stage2"
        oTable.Cell(row, 2).Range.Text = TextBox79.Text
        oTable.Cell(row, 3).Range.Text = "[Pa]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Efficiency stage2"
        oTable.Cell(row, 2).Range.Text = TextBox109.Text
        oTable.Cell(row, 3).Range.Text = "[%]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage2"
        oTable.Cell(row, 2).Range.Text = TextBox108.Text
        oTable.Cell(row, 3).Range.Text = "[mg/Am3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage2"
        oTable.Cell(row, 2).Range.Text = TextBox179.Text
        oTable.Cell(row, 3).Range.Text = "[mg/Nm3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage2"
        oTable.Cell(row, 2).Range.Text = TextBox184.Text
        oTable.Cell(row, 3).Range.Text = "[kg/h]"
        row += 2
        oTable.Cell(row, 1).Range.Text = "Efficiency stage1+2"
        oTable.Cell(row, 2).Range.Text = TextBox120.Text
        oTable.Cell(row, 3).Range.Text = "[%]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage1+2"
        oTable.Cell(row, 2).Range.Text = TextBox134.Text
        oTable.Cell(row, 3).Range.Text = "[mg/Am3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage1+2"
        oTable.Cell(row, 2).Range.Text = TextBox178.Text
        oTable.Cell(row, 3).Range.Text = "[mg/Nm3]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "Emission stage1+2"
        oTable.Cell(row, 2).Range.Text = TextBox184.Text
        oTable.Cell(row, 3).Range.Text = "[kg/h]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "dp(100) stage1+2"
        oTable.Cell(row, 2).Range.Text = TextBox119.Text
        oTable.Cell(row, 3).Range.Text = "[mu] (100% loss)"
        row += 1
        oTable.Cell(row, 1).Range.Text = "dp(50) stage1+2"
        oTable.Cell(row, 2).Range.Text = TextBox123.Text
        oTable.Cell(row, 3).Range.Text = "[mu]"
        row += 1
        oTable.Cell(row, 1).Range.Text = "dp(5) stage1+2"
        oTable.Cell(row, 2).Range.Text = TextBox125.Text
        oTable.Cell(row, 3).Range.Text = "[mu]"

        oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns 
        oTable.Columns(2).Width = oWord.InchesToPoints(1)
        oTable.Columns(3).Width = oWord.InchesToPoints(2)

        oTable.Rows(1).Range.Font.Bold = CInt(True)
        oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

        '------------------save Chart2 (Loss curve)---------------- 
        Draw_chart2(Chart2, 0)
        file_name = dirpath_Temp & "Chart_loss.Jpeg"
        'Chart2.SaveImage(file_name, System.Drawing.Imaging.ImageFormat.Jpeg)
        Chart1.SaveImage(file_name, System.Drawing.Imaging.ImageFormat.Jpeg)
        oPara4 = oDoc.Content.Paragraphs.Add
        oPara4.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
        oPara4.Range.InlineShapes.AddPicture(file_name)
        oPara4.Range.InlineShapes.Item(1).LockAspectRatio = CType(True, Microsoft.Office.Core.MsoTriState)
        oPara4.Range.InlineShapes.Item(1).ScaleWidth = chart_size       'Size
        oPara4.Range.InsertParagraphAfter()
    End Sub
    'Air viscosity
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown21.ValueChanged
        Dim Visco As Double
        Visco = Air_visco(NumericUpDown21.Value)
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
        Return (vis * 100.0)    '[kg/m-s]-->[centi Poise]
    End Function
    Private Sub Present_loss_grid1(ks As Integer)
        Dim j As Integer

        DataGridView2.ColumnCount = 18
        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add(111)
        DataGridView2.RowHeadersVisible = False

        DataGridView2.EnableHeadersVisualStyles = False                     'For backcolor
        DataGridView2.Columns(0).HeaderText = "Dia upper [mu]"              '
        DataGridView2.Columns(1).HeaderText = "Dia (aver.) Class [mu]"      '         
        DataGridView2.Columns(2).HeaderText = "Dia/ k_stokes [-]"            '
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

        'DataGridView2.AutoSize = False
        For row = 1 To 110  'Fill the DataGrid
            j = row - 1
            DataGridView2.Rows(j).Cells(0).Value = _cees(ks).stage1(j).dia.ToString("F8")
            DataGridView2.Rows(j).Cells(1).Value = _cees(ks).stage1(j).d_ave.ToString("F8")        'Average diameter
            DataGridView2.Rows(j).Cells(2).Value = _cees(ks).stage1(j).d_ave_K.ToString("F5")      'Average dia/K stokes
            DataGridView2.Rows(j).Cells(3).Value = _cees(ks).stage1(j).loss_overall.ToString("F5")      'Loss 
            DataGridView2.Rows(j).Cells(4).Value = _cees(ks).stage1(j).loss_overall_C.ToString("F5")    'Loss 
            DataGridView2.Rows(j).Cells(5).Value = _cees(ks).stage1(j).catch_chart.ToString("F5")       'Catch
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
        DataGridView2.Rows(111).Cells(15).Value = _cees(ks).sum_psd_diff1.ToString("F5")    'total_psd_diff.
        DataGridView2.Rows(111).Cells(16).Value = _cees(ks).sum_loss1.ToString("F5")        'total_abs_loss.ToString("F5")
        DataGridView2.Rows(111).Cells(17).Value = _cees(ks).sum_loss_C1.ToString("F5")      'total_abs_loss_C.ToString("F5")
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    End Sub

    Private Sub Present_loss_grid2(ks As Integer)
        Dim j As Integer

        DataGridView3.ColumnCount = 19
        DataGridView3.Rows.Clear()
        DataGridView3.Rows.Add(111)
        DataGridView3.RowHeadersVisible = False

        DataGridView3.EnableHeadersVisualStyles = False                         'For backcolor
        DataGridView3.Columns(0).HeaderText = "Dia upper [mu]"                  '
        DataGridView3.Columns(1).HeaderText = "Dia (aver.) Class [mu]"          '
        DataGridView3.Columns(2).HeaderText = "Dia/ k_stokes [-]"                       '
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
        DataGridView3.Columns(14).HeaderText = "psd_cum_pro For chart"          '
        DataGridView3.Columns(15).HeaderText = "psd diff [%] Of 1-stage"        '
        DataGridView3.Columns(16).HeaderText = "psd diff [%] Of 2-stage"        '
        DataGridView3.Columns(17).HeaderText = "catch loss abs [%]"             '
        DataGridView3.Columns(18).HeaderText = "loss abs corrected [%]"         '

        DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

        DataGridView3.Columns(1).HeaderCell.Style.BackColor = Color.Yellow      'For chart
        DataGridView3.Columns(5).HeaderCell.Style.BackColor = Color.Yellow      'For chart
        DataGridView3.Columns(14).HeaderCell.Style.BackColor = Color.Yellow     'For chart

        For row = 1 To 110  'Fill the DataGrid
            j = row - 1
            DataGridView3.Rows(j).Cells(0).Value = _cees(ks).stage2(j).dia.ToString("F8")         'Dia particle
            DataGridView3.Rows(j).Cells(1).Value = _cees(ks).stage2(j).d_ave.ToString("F8")        'Average diameter
            DataGridView3.Rows(j).Cells(2).Value = _cees(ks).stage2(j).d_ave_K.ToString("F5")      'Average dia/K stokes
            DataGridView3.Rows(j).Cells(3).Value = _cees(ks).stage2(j).loss_overall.ToString("F5") 'Loss 
            DataGridView3.Rows(j).Cells(4).Value = _cees(ks).stage2(j).loss_overall_C.ToString("F5") 'Loss corrected 
            DataGridView3.Rows(j).Cells(5).Value = _cees(ks).stage2(j).catch_chart.ToString("F5")  'Catch for chart
            DataGridView3.Rows(j).Cells(6).Value = _cees(ks).stage2(j).i_grp.ToString             'Groep nummer
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
        If Double.IsNaN(k) Then k = 1.0
        If Double.IsNaN(m) Then m = 1.0
        'Debug.WriteLine("")
        'Debug.WriteLine("g.i_p1= " & g.i_p1.ToString & ", g.i_p2= " & g.i_p2.ToString)
        'Debug.WriteLine("g.i_d1= " & g.i_d1.ToString & ", g.i_d2= " & g.i_d2.ToString)
        'Debug.WriteLine("k= " & k.ToString & ", g.m= " & m.ToString)

        g.i_k = k
        g.i_m = m
    End Sub

    Private Sub Calc_stage1(ks As Integer)
        'This is the standard VTK cyclone calculation for case "ks" 
        Dim i As Integer = 0
        Dim perc_smallest_part1 As Double   '[%]
        Dim fac_m As Double                 '[-]
        Dim words As String()
        Dim density_ratio1 As Double        '[-] Density ratio stage#1  kg/Nm3 and kg/Am3
        Dim density_ratio2 As Double        '[-] Density ratio stage#1  kg/Nm3 and kg/Am3
        Dim dc1 As Double                   '[m] diameter cyclone stage #1
        Dim dc2 As Double                   '[m] diameter cyclone stage #2
        Dim dustload As Double              '[g/Am3] dust load

        If Double.IsNaN(_cees(ks).stage1(0).dia) OrElse Double.IsInfinity(_cees(ks).stage1(0).dia) Then
            MessageBox.Show("stage1 diameter1 = indifinity")
        End If

        If ComboBox1.Items.Count > 0 AndAlso ComboBox2.Items.Count > 0 AndAlso init Then

            dustload = NumericUpDown4.Value     '[g/Am3] dust load

            '------ the idea is that the smallest diameter cyclone determines
            '------ the smallest particle diameter used in the calculation
            '------ for the stage #1 cyclone
            dc1 = numericUpDown5.Value         '[m] Diameter cyclone stage 1
            dc2 = NumericUpDown34.Value        '[m] Diameter cyclone stage 2
            If dc1 > dc2 Then
                _cees(ks).stage1(0).dia = Calc_dia_particle(1.0, _cees(ks).Kstokes2, 2) 'stage #2 cyclone
            Else
                _cees(ks).stage1(0).dia = Calc_dia_particle(1.0, _cees(ks).Kstokes1, 1) 'stage #1 cyclone
            End If

            _cees(ks).stage1(0).d_ave = _cees(ks).stage1(0).dia / 2.0                      'Average diameter
            _cees(ks).stage1(0).d_ave_K = _cees(ks).stage1(0).d_ave / _cees(ks).Kstokes1  'dia/k_stokes
            _cees(ks).stage1(0).loss_overall = Calc_verlies(_cees(ks).stage1(0).d_ave_K, _cees(ks).Kstokes1, 1)     '[-] loss overall
            Calc_verlies_corrected(_cees(ks).stage1(0), 1)                                '[-] loss overall corrected
            _cees(ks).stage1(0).catch_chart = (1.0 - _cees(ks).stage1(i).loss_overall_C) * 100.0     '[%]

            Calc_diam_classification(_cees(ks).stage1(0))        'Classify this part size
            Calc_k_and_m(_cees(ks).stage1(0))                        'Calculate i_m and i_k

            _cees(ks).stage1(0).psd_cum = Math.E ^ (-((_cees(ks).stage1(i).dia / _cees(ks).stage1(i).i_m) ^ _cees(ks).stage1(i).i_k))
            _cees(ks).stage1(0).psd_cum_pro = _cees(ks).stage1(i).psd_cum * 100

            _cees(ks).stage1(0).psd_dif = 100.0 * (1.0 - _cees(ks).stage1(i).psd_cum)
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
            If ComboBox1.Items.Count > 0 Then    'Prevent start up problems
                words = rekenlijnen(ComboBox1.SelectedIndex).Split(CType(";", Char()))
                '---- diameter kleiner dan dia kritisch
                fac_m = CDbl(words(2))
            End If

            perc_smallest_part1 = 0.0000001                      'smallest particle [%]
            _cees(ks).Dmax1 = Calc_dia_particle(perc_smallest_part1, _cees(ks).Kstokes1, 1)     '=100% loss (biggest particle)
            _cees(ks).Dmin1 = _cees(ks).Kstokes1 * fac_m        'diameter smallest particle caught by this type cyclone

            ' TextBox24.Text &= "_cees(ks).Kstokes1= " & _cees(ks).Kstokes1.ToString & ",  fac_m= " & fac_m.ToString & ",  _cees(ks).Dmin1= " & _cees(ks).Dmin1.ToString & vbCrLf

            '------------ Particle diameter calculation step -----
            '====== _cees(ks).Dmin2 === is the result fram calculation stage 2 ===
            '====== and maybe unknown ============================================
            '============ With start iteration the Dmin2 is unknown =========

            If _cees(ks).Dmin2 < 0.05 Then _cees(ks).Dmin2 = 0.05
            _istep = (_cees(ks).Dmax1 / _cees(ks).Dmin2) ^ (1.0 / 110.0) 'Calculation step

            'TextBox24.Text &= "_istep= " & _istep.ToString & ",  _cees(ks).Dmax1= " & _cees(ks).Dmax1.ToString & vbCrLf

            For i = 1 To 110
                _cees(ks).stage1(i).dia = _cees(ks).stage1(i - 1).dia * _istep
                _cees(ks).stage1(i).d_ave = (_cees(ks).stage1(i - 1).dia + _cees(ks).stage1(i).dia) / 2   'Average diameter
                _cees(ks).stage1(i).d_ave_K = _cees(ks).stage1(i).d_ave / _cees(ks).Kstokes1                'dia/k_stokes
                _cees(ks).stage1(i).loss_overall = Calc_verlies(_cees(ks).stage1(i).d_ave, _cees(ks).Kstokes1, 1)   '[-] loss overall
                Calc_verlies_corrected(_cees(ks).stage1(i), 1)                                              '[-] loss overall corrected

                If CheckBox2.Checked Then
                    _cees(ks).stage1(i).catch_chart = (1 - _cees(ks).stage1(i).loss_overall_C) * 100.0  '[%] Corrected
                Else
                    _cees(ks).stage1(i).catch_chart = (1 - _cees(ks).stage1(i).loss_overall) * 100.0    '[%] NOT corrected
                End If

                Calc_diam_classification(_cees(ks).stage1(i))         'Classify this part size (result is  .i_grp)

                '====to prevent silly results====
                If _cees(ks).stage1(i).i_grp < no_PDS_inputs Then  'OK in this one counts
                    Calc_k_and_m(_cees(ks).stage1(i))       'Calculate i_m and i_k (based on particle dia. and percentages)
                    _cees(ks).stage1(i).psd_cum = Math.E ^ (-((_cees(ks).stage1(i).dia / _cees(ks).stage1(i).i_m) ^ _cees(ks).stage1(i).i_k))
                    _cees(ks).stage1(i).psd_cum_pro = _cees(ks).stage1(i).psd_cum * 100
                    _cees(ks).stage1(i).psd_dif = 100.0 * (_cees(ks).stage1(i - 1).psd_cum - _cees(ks).stage1(i).psd_cum)
                Else
                    _cees(ks).stage1(i).i_k = 0.0
                    _cees(ks).stage1(i).i_m = 0.0
                    _cees(ks).stage1(i).psd_cum = 0.0
                    _cees(ks).stage1(i).psd_cum_pro = 0.0
                    _cees(ks).stage1(i).psd_dif = 0.0
                    _cees(ks).stage1(i).i_grp = 0.0
                End If

                _cees(ks).stage1(i).loss_abs = _cees(ks).stage1(i).loss_overall * _cees(ks).stage1(i).psd_dif
                _cees(ks).stage1(i).loss_abs_C = _cees(ks).stage1(i).loss_overall_C * _cees(ks).stage1(i).psd_dif

                '----- sum value incremental values -----
                _cees(ks).sum_psd_diff1 += _cees(ks).stage1(i).psd_dif
                _cees(ks).sum_loss1 += _cees(ks).stage1(i).loss_abs         '[%] Summ loss
                _cees(ks).sum_loss_C1 += _cees(ks).stage1(i).loss_abs_C     '[%] Summ loss corrected
            Next

            _cees(ks).loss_total1 = _cees(ks).sum_loss_C1 + ((100.0 - _cees(ks).sum_psd_diff1) * perc_smallest_part1)

            _cees(ks).emmis1_Am3 = dustload * (_cees(ks).loss_total1 / 100.0)  '[g/Am3]
            _cees(ks).emmis1_Nm3 = _cees(ks).emmis1_Am3 * Calc_Normal_density(_cees(ks).Ro_gas1_Am3, _cees(ks).p1_abs, _cees(ks).Temp)

            '----------Dust load stage #2 is emission stage #1 -----------
            _cees(ks).dust2_Am3 = _cees(ks).emmis1_Am3

            '--------- Density ratio stage #2 kg/Nm3 and kg/Am3 ---------
            density_ratio2 = _cees(ks).Ro_gas2_Nm3 / _cees(ks).Ro_gas2_Am3  '[-]
            _cees(ks).dust2_Nm3 = _cees(ks).dust2_Am3 * density_ratio2      'Dust load [gram/Nm3]

            CheckBox3.Checked = CBool(IIf(_cees(ks).dust2_Am3 > 20.0, True, False))
            _cees(ks).Efficiency1 = 100.0 - _cees(ks).loss_total1        '[%] Efficiency

            '----------- present stage #1-----------
            TextBox51.Text = _cees(ks).Dmax1.ToString("F2")     'diameter [mu] 100% catch
            TextBox52.Text = _cees(ks).Dmin1.ToString("F2")     'diameter [mu] 100% loss
            TextBox56.Text = ComboBox1.Text                     'Cyclone type
            TextBox57.Text = CheckBox2.Checked.ToString         'Correction stage #1
            TextBox70.Text = _cees(ks).dust2_Am3.ToString("F3") 'Dust load [gram/Am3]

            TextBox118.Text = _cees(ks).sum_psd_diff1.ToString("F3")
            TextBox54.Text = _cees(ks).sum_loss1.ToString("F3")     '[%] Summ loss
            TextBox34.Text = _cees(ks).sum_loss_C1.ToString("F3")   '[%] Summ loss Corrected

            '---------Density ratio stage#1  kg/Nm3 and kg/Am3 ---------
            density_ratio1 = _cees(ks).Ro_gas1_Nm3 / _cees(ks).Ro_gas1_Am3  '[-]
            _cees(ks).emmis1_Nm3 = _cees(ks).emmis1_Am3 * density_ratio1    '[g/Am3]
            _cees(ks).emis1_kgh = _cees(ks).dust1_in_kgh * (100.0 - _cees(ks).Efficiency1) / 100.0


            '----------- Dust load correction stage #1 ------------------
            If CheckBox2.Checked Then 'High load
                TextBox58.Text = TextBox34.Text                             '[%] Corrected 
            Else
                TextBox58.Text = TextBox54.Text                             '[%] NOT Corrected
            End If

            '---------- present ------------------
            TextBox59.Text = _cees(ks).Efficiency1.ToString("F2")            '[%]
            TextBox21.Text = _cees(ks).Efficiency1.ToString("F3")            '[%]
            TextBox60.Text = (_cees(ks).emmis1_Am3 * 1000).ToString("F1")    '[mg/Am3]

            '==== 1st stage ====
            TextBox18.Text = (_cees(ks).emmis1_Am3 * 1000).ToString("F1")    '[mg/Am3]
            TextBox133.Text = (_cees(ks).emmis1_Nm3 * 1000).ToString("F1")   '[mg/Nm3]
            TextBox180.Text = (_cees(ks).emmis1_Nm3 * 1000).ToString("F1")   '[mg/Nm3]
            TextBox186.Text = _cees(ks).emis1_kgh.ToString("F2")             '[kg/h]
        End If
    End Sub

    Private Sub Calc_stage2(ks As Integer)
        'This is the standard VTK cyclone calculation 
        Dim i As Integer = 0
        Dim perc_smallest_part2 As Double
        Dim fac_m As Double
        Dim words As String()
        Dim kgh, tot_kgh As Double
        Dim Eff_comb As Double      'Efficiency stage #1 and #2

        If Double.IsNaN(_cees(ks).stage1(0).dia) OrElse Double.IsInfinity(_cees(ks).stage1(0).dia) Then
            MessageBox.Show("stage2 diameter = indifinity")
        End If

        If ComboBox1.Items.Count > 0 AndAlso ComboBox2.Items.Count > 0 AndAlso init Then

            '----------- stof belasting ------------
            tot_kgh = _cees(ks).emis1_kgh                       '[kg/hr] Dust inlet 
            kgh = tot_kgh / _cees(ks).Noc2                      '[kg/hr/Cy] Dust inlet 

            TextBox100.Text = kgh.ToString("F2")                '[kg/hr] Dust inlet 
            TextBox101.Text = tot_kgh.ToString("F2")            '[kg/hr/Cy] Dust inlet 

            '--------- now the particles (====Grid line 0======)------------
            _cees(ks).stage2(0).dia = _cees(ks).stage1(0).dia                                   'Copy stage #1
            _cees(ks).stage2(0).d_ave = _cees(ks).stage2(0).dia / 2.0                           'Average diameter
            _cees(ks).stage2(0).d_ave_K = _cees(ks).stage2(0).d_ave / _cees(ks).Kstokes2        'dia/k_stokes
            _cees(ks).stage2(0).loss_overall = Calc_verlies(_cees(ks).stage2(0).d_ave_K, _cees(ks).Kstokes2, 2)     '[-] loss overall
            Calc_verlies_corrected(_cees(ks).stage2(0), 2)                                      '[-] loss overall corrected
            _cees(ks).stage2(0).catch_chart = (1.0 - _cees(ks).stage2(0).loss_overall_C) * 100.0  '[%]
            Calc_diam_classification(_cees(ks).stage2(0))                                       'groepnummer

            If _cees(ks).stage2(i).i_grp < no_PDS_inputs Then
                Calc_k_and_m(_cees(ks).stage2(0))
                _cees(ks).stage2(0).psd_cum = Math.E ^ (-((_cees(ks).stage2(0).dia / _cees(ks).stage2(0).i_m) ^ _cees(ks).stage2(0).i_k))
                _cees(ks).stage2(0).psd_cum_pro = _cees(ks).stage2(0).psd_cum * 100.0 '[%]
                _cees(ks).stage2(0).psd_dif = 100.0 * _cees(ks).stage1(0).loss_abs / (100.0 - _cees(ks).Efficiency1)              'LOSS STAGE #1
            Else
                _cees(ks).stage2(0).i_k = 0.0
                _cees(ks).stage2(0).i_m = 0.0
                _cees(ks).stage2(0).psd_cum = 0.0
                _cees(ks).stage2(0).psd_cum_pro = 0.0
                _cees(ks).stage2(0).psd_dif = 0.0
            End If

            _cees(ks).stage2(0).loss_abs = _cees(ks).stage2(0).loss_overall * _cees(ks).stage2(0).psd_dif
            _cees(ks).stage2(0).loss_abs_C = _cees(ks).stage2(0).loss_overall_C * _cees(ks).stage2(0).psd_dif

            '----- initial values -------
            _cees(ks).sum_psd_diff2 = 0.0       '_cees(ks).stage2(0).psd_dif
            _cees(ks).sum_loss2 = 0.0           '_cees(ks).stage2(0).loss_abs
            _cees(ks).sum_loss_C2 = 0.0         '_cees(ks).stage2(0).loss_abs_C

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
                If Not (Double.IsNaN(_cees(ks).stage1(ks).dia) OrElse Double.IsInfinity(_cees(ks).stage1(ks).dia)) Then

                    _cees(ks).stage2(i).dia = _cees(ks).stage1(i).dia                               'Diameter Copy stage #1
                    _cees(ks).stage2(i).d_ave = _cees(ks).stage1(i).d_ave                           'Average diameter
                    _cees(ks).stage2(i).d_ave_K = _cees(ks).stage2(i).d_ave / _cees(ks).Kstokes2    'dia/k_stokes
                    _cees(ks).stage2(i).loss_overall = Calc_verlies(_cees(ks).stage2(i).d_ave, _cees(ks).Kstokes2, 2)   '[-] loss overall
                    Calc_verlies_corrected(_cees(ks).stage2(i), 2)                                  '[-] loss overall corrected

                    '------------- Load correction stage #2 -------------------
                    If CheckBox3.Checked Then
                        _cees(ks).stage2(i).catch_chart = (1 - _cees(ks).stage2(i).loss_overall_C) * 100.0  '[%] Corrected
                    Else
                        _cees(ks).stage2(i).catch_chart = (1 - _cees(ks).stage2(i).loss_overall) * 100.0    '[%] NOT corrected
                    End If

                    Calc_diam_classification(_cees(ks).stage2(i))                                     'Calc
                    _cees(ks).stage2(i).i_grp = _cees(ks).stage1(i).i_grp

                    If _cees(ks).stage2(i).i_grp < no_PDS_inputs Then
                        Calc_k_and_m(_cees(ks).stage2(i))
                        _cees(ks).stage2(i).psd_cum = Math.E ^ (-(_cees(ks).stage2(i).dia / _cees(ks).stage2(i).i_m) ^ _cees(ks).stage2(i).i_k)
                        _cees(ks).stage2(i).psd_cum_pro = _cees(ks).stage2(i).psd_cum * 100.0 '[%]
                        _cees(ks).stage2(i).psd_dif = 100.0 * _cees(ks).stage1(i).loss_abs_C / _cees(ks).sum_loss_C1
                    Else
                        _cees(ks).stage2(i).i_k = 0.0
                        _cees(ks).stage2(i).i_m = 0.0
                        _cees(ks).stage2(i).psd_cum = 0.0
                        _cees(ks).stage2(i).psd_cum_pro = 0.0
                        _cees(ks).stage2(i).psd_dif = 0.0
                    End If

                    _cees(ks).stage2(i).loss_abs = _cees(ks).stage2(i).loss_overall * _cees(ks).stage2(i).psd_dif
                    _cees(ks).stage2(i).loss_abs_C = _cees(ks).stage2(i).loss_overall_C * _cees(ks).stage2(i).psd_dif

                    '----- sum value incremental values -----
                    _cees(ks).sum_psd_diff2 += _cees(ks).stage2(i).psd_dif
                    _cees(ks).sum_loss2 += _cees(ks).stage2(i).loss_abs
                    _cees(ks).sum_loss_C2 += _cees(ks).stage2(i).loss_abs_C
                End If
            Next i
            _cees(ks).loss_total2 = _cees(ks).sum_loss_C2 + ((100.0 - _cees(ks).sum_psd_diff2) * perc_smallest_part2)
            _cees(ks).emmis2_Am3 = _cees(ks).emmis1_Am3 * _cees(ks).loss_total2 / 100.0
            _cees(ks).Efficiency2 = 100.0 - _cees(ks).loss_total2      '[%] Efficiency

            '---------Density ratio stage#2  kg/Nm3 and kg/Am3 ---------
            Dim density_ratio2 As Double
            density_ratio2 = _cees(ks).Ro_gas2_Nm3 / _cees(ks).Ro_gas2_Am3  '[-]
            _cees(ks).emmis2_Nm3 = _cees(ks).emmis2_Am3 * density_ratio2    '[g/Am3]

            '------ combined efficiency -----
            Eff_comb = _cees(ks).Efficiency1 + (1 - _cees(ks).Efficiency1 / 100.0) * _cees(ks).Efficiency2
            _cees(ks).emis2_kgh = _cees(ks).dust1_in_kgh * (100.0 - Eff_comb) / 100.0

            '----------- present stage #2 -----------
            TextBox63.Text = ComboBox2.Text                     'Cyclone type
            TextBox64.Text = CheckBox3.Checked.ToString         'Hi load correction stage #2
            TextBox110.Text = _cees(ks).Dmax2.ToString("F2")    'diameter [mu] 100% catch
            TextBox111.Text = _cees(ks).Dmin2.ToString("F2")    'diameter [mu] 100% loss
            TextBox116.Text = _istep.ToString("F3")             'Calculation step stage #2
            TextBox43.Text = _cees(ks).Dmin1.ToString("F2")     'Smallest part caught by the second stage cyclone

            TextBox117.Text = _cees(ks).sum_psd_diff2.ToString("F3")
            TextBox68.Text = _cees(ks).sum_loss2.ToString("F3")
            TextBox69.Text = _cees(ks).sum_loss_C2.ToString("F3")
            TextBox120.Text = Eff_comb.ToString("F4")

            If CheckBox3.Checked Then   'Dust load correction stage #2
                TextBox65.Text = _cees(ks).loss_total2.ToString("F3")    '[%] Corrected
            Else
                TextBox65.Text = _cees(ks).sum_loss2.ToString("F3")      '[%] NOT Corrected
            End If

            TextBox66.Text = _cees(ks).Efficiency2.ToString("F2")       '[%] stage #2
            TextBox109.Text = _cees(ks).Efficiency2.ToString("F3")      '[%]
            TextBox62.Text = (_cees(ks).emmis2_Am3 * 1000).ToString("F1")       '[mgram/Am3] emmissie stage #2
            TextBox108.Text = (_cees(ks).emmis2_Am3 * 1000).ToString("F1")      '[mgram/Am3] emmissie stage #2
            TextBox134.Text = (_cees(ks).emmis2_Am3 * 1000).ToString("F1")      '[mgram/Am3] emmissie stage #2

            TextBox142.Text = (_cees(ks).emmis2_Nm3 * 1000).ToString("F1")      '[mgram/Nm3] emmissie stage #2
            TextBox179.Text = (_cees(ks).emmis2_Nm3 * 1000).ToString("F1")      '[mgram/Nm3] emmissie stage #2
            TextBox178.Text = (_cees(ks).emmis2_Nm3 * 1000).ToString("F1")      '[mgram/Nm3] emmissie stage #2

            TextBox185.Text = _cees(ks).emis2_kgh.ToString("F2")                '[kg/h] emmissie stage #2
            TextBox184.Text = _cees(ks).emis2_kgh.ToString("F2")                '[kg/h] emmissie stage #2

            TextBox143.Text = (_cees(ks).dust2_Nm3 * 1000).ToString("F2")        'Dust load [mgram/Nm3]
        End If
    End Sub

    Public Sub Calc_diam_classification(ByRef g As GvG_Calc_struct)
        'Determine the particle diameter class 
        Dim grp_count As Integer = 0

        '=========== Determine how many PSD groups are there? ============
        For i = 0 To no_PDS_inputs - 1
            If (_input(i).dia_big > 0) AndAlso (_input(i).class_load > 0) Then grp_count += 1
        Next
        g.i_grp = 0

        If g.dia >= 0 Then
            '=========== first entered data point ===========
            If (g.dia < _input(0).dia_big) Then
                g.i_d1 = _input(0).dia_big     'Diameter small [mu]
                g.i_d2 = _input(1).dia_big     'Diameter big [mu]
                g.i_p1 = _input(0).class_load  'User lower input percentage
                g.i_p2 = _input(1).class_load  'User upper input percentage
                g.i_grp = 0                    'Group 0
            End If

            '=========== mid section ===========
            For i = 1 To grp_count
                If (g.dia >= _input(i - 1).dia_big AndAlso g.dia < _input(i).dia_big AndAlso _input(i).dia_big > 0) Then
                    g.i_d1 = _input(i - 1).dia_big              'Diameter small [mu]
                    g.i_d2 = _input(i).dia_big                  'Diameter big [mu]
                    g.i_p1 = _input(i - 1).class_load           'User lower input percentage
                    g.i_p2 = _input(i).class_load               'User upper input percentage
                    g.i_grp = i                                 'Group 1 up and including 11
                End If
            Next

            '=========== last entered PSD data point ===========
            If (g.dia >= _input(no_PDS_inputs).dia_big AndAlso _input(no_PDS_inputs).dia_big > 0) Then
                g.i_d1 = _input(no_PDS_inputs).dia_big        'Diameter small [mu]
                g.i_d2 = 2000                                       'Diameter big [mu]
                g.i_p1 = _input(no_PDS_inputs).class_load           'User lower input percentage
                g.i_p2 = 0                                          'User upper input percentage
                g.i_grp = grp_count                                 'Last PSD input star
            End If

            Dim w(no_PDS_inputs) As Double    'Individual particle class weights 
            'Dim q(no_PDS_inputs) As Double    'Individual particle class weights 
            Dim qsum(no_PDS_inputs) As Double 'Sum of weights
            Dim j As Integer

            qsum(0) = 0
            For i = 1 To no_PDS_inputs - 1
                qsum(i) = qsum(i - 1) - w(i - 1)
            Next

            For i = 0 To no_PDS_inputs - 1
                j = (w.Length - 1 - i)
                w(i) = _input(j).class_load - Abs(qsum(i))
            Next

            Label1.Visible = False  'Error message
        Else
            MessageBox.Show("Error In Calc_diam_classification")
            Label1.Visible = True   'Error message
        End If
    End Sub
    Private Sub Check_DGV6()

        With DataGridView6
            If .RowCount > 0 Then
                Dim dia, dia_previous As Double
                Dim c_load, c_load_previous As Double

                '---------CHECK- diameters must increase-----------
                .Rows(0).Cells(0).Style.BackColor = Color.LightGreen
                For i = 1 To .Rows.Count - 1
                    dia = _input(i).dia_big
                    dia_previous = _input(i - 1).dia_big

                    .Rows(i).Cells(0).Style.BackColor = CType(IIf(dia < dia_previous AndAlso dia <> 0, Color.Red, Color.LightGreen), Color)
                Next

                '---------CHECK-cummulative weight must decrease-----------
                .Rows(0).Cells(1).Style.BackColor = Color.LightGreen
                For i = 1 To .Rows.Count - 1
                    c_load = _input(i).class_load
                    c_load_previous = _input(i - 1).class_load

                    ' .Rows(i).Cells(1).Style.BackColor = If(c_load > c_load_previous And c_load <> 0, Color.Orange, Color.LightGreen)
                    If (c_load > c_load_previous AndAlso c_load <> 0) Then
                        .Rows(i).Cells(1).Style.BackColor = Color.Orange
                    Else
                        .Rows(i).Cells(1).Style.BackColor = Color.LightGreen
                    End If
                Next

                '--- Value 100 gives errors -----
                Dim qq As Double = CDbl(.Rows(0).Cells(1).Value)
                If qq >= 100.0 Then
                    .Rows(0).Cells(1).Value = "99.9999"
                End If
            End If
        End With

    End Sub

    'Calculate cyclone weight
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click, TabPage7.Enter, NumericUpDown32.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown45.ValueChanged, NumericUpDown44.ValueChanged, NumericUpDown42.ValueChanged, NumericUpDown41.ValueChanged
        Calc_cycl_weight()
        Calc_Tang_outlet()
    End Sub

    Private Sub Calc_cycl_weight()
        Dim wc1(6) As Double                '[kg] cyclone1 parts weight       
        Dim wc2(6) As Double                '[kg] cyclone2 parts weight
        Dim area_c1(6) As Double            '[m2] cyclone1 parts weight       
        Dim area_c2(6) As Double            '[m2] cyclone2 parts weight

        Dim c_weight1 As Double             '[kg] cyclone weight
        Dim c_weight2 As Double             '[kg] cyclone weight
        Dim total_instal As Double          '[kg] installation weight

        Dim _db1 As Double = numericUpDown5.Value                 '[m] Body diameter
        Dim plt_top1 As Double = CDbl(NumericUpDown32.Value) / 1000     '[m] top plate
        Dim plt_top2 As Double = CDbl(NumericUpDown41.Value) / 1000     '[m] top plate

        Dim _db2 As Double = NumericUpDown34.Value                '[m] Body diameter
        Dim plt_body1 As Double = CDbl(NumericUpDown31.Value) / 1000    '[m] rest of the cyclone
        Dim plt_body2 As Double = CDbl(NumericUpDown42.Value) / 1000    '[m] rest of the cyclone

        Dim ro_steel As Double = 7850.0     '[kg/m3] Density steel
        Dim hh, hj, hk As Double            'Dimensions
        Dim sheet_metal_wht As Double
        Dim p304, p316 As Double
        Dim c_area1, c_area2 As Double      'Outside area est.

        If _db1 > 0.001 AndAlso _db2 > 0.001 Then

            '========== Stage cyclone #1 =====
            'weight top plate
            area_c1(0) = PI / 4 * _db1 ^ 2 * 1.2            '[m2] (not accurate)
            wc1(0) = area_c1(0) * plt_top1 * ro_steel       '[kg]

            'weight cylindrical body
            hh = _cyl1_dim(10) * _db1                      '[m] Length romp
            area_c1(1) = PI * _db1 * hh                    '[m2] romp
            wc1(1) = area_c1(1) * plt_body1 * ro_steel     '[kg] weight romp

            'weight cone
            hh = _cyl1_dim(11) * _db1                       '[m] Length cone
            hj = _db1                                       '[m] grote diameter cone 
            hk = _cyl1_dim(12) * _db1                       '[m] kleine diameter cone 
            area_c1(2) = PI * (hj + hk) / 2 * hh            '[m2] cone
            wc1(2) = area_c1(2) * plt_body1 * ro_steel      '[kg] weight cone

            'weight gas outlet pipe
            hh = _cyl1_dim(8) * _db1                        '[m] Length insteekpijp
            hj = _cyl1_dim(7) * _db1                        '[m] Uitlaat flensdiameter inw.
            area_c1(3) = PI * hh * hj
            wc1(3) = area_c1(3) * plt_body1 * ro_steel    '[kg] weight insteekpijp


            'weight 3P pipe
            hh = _cyl1_dim(13) * _db1                       '[m] Length 3P pijp
            hj = _cyl1_dim(12) * _db1                       '[m] diameter 3P inw.
            area_c1(4) = PI * hj * hh
            wc1(4) = area_c1(4) * plt_body1 * ro_steel    '[kg] weight 3P pipe

            'weight 3P cone
            hh = _cyl1_dim(14) * _db1                       '[m] Length 3P cone
            hj = _cyl1_dim(12) * _db1                       '[m] grote diameter 3P pijp
            hk = _cyl1_dim(15) * _db1                       '[m] kleine diameter 3P pijp
            area_c1(5) = PI * (hj + hk) / 2 * hh
            wc1(5) = area_c1(5) * plt_body1 * ro_steel '[kg] weight 3P pipe

            For i = 0 To wc1.Length - 1
                c_weight1 += wc1(i)                          'Total weight
                c_area1 += area_c1(i)
            Next

            c_weight1 *= 1.1                                '10% weight flanges +safety


            '========== Stage cyclone #2 ========
            'weight top plate
            area_c2(0) = PI / 4 * _db1 ^ 2 * 1.2             '[m2] (not accurate)
            wc2(0) = area_c2(0) * plt_top2 * ro_steel        '[kg]

            'weight cylindrical body
            hh = _cyl1_dim(10) * _db2                       '[m] Length romp
            area_c2(1) = PI * _db2 * hh                     '[m2] romp
            wc2(1) = area_c2(1) * plt_body2 * ro_steel      '[kg] weight romp

            'weight cone
            hh = _cyl1_dim(11) * _db2                       '[m] Length cone
            hj = _db2                                       '[m] grote diameter cone 
            hk = _cyl1_dim(12) * _db2                       '[m] kleine diameter cone 
            area_c2(2) = PI * (hj + hk) / 2 * hh            '[m2] cone
            wc2(2) = area_c2(2) * plt_body2 * ro_steel      '[kg] weight cone

            'weight gas outlet pipe
            hh = _cyl1_dim(8) * _db2                        '[m] Length insteekpijp
            hj = _cyl1_dim(7) * _db2                        '[m] Uitlaat flensdiameter inw.
            area_c2(3) = PI * hh * hj
            wc2(3) = area_c2(3) * plt_body2 * ro_steel      '[kg] weight insteekpijp

            'weight 3P pipe
            hh = _cyl1_dim(13) * _db2                       '[m] Length 3P pijp
            hj = _cyl1_dim(12) * _db2                       '[m] diameter 3P inw.
            area_c2(4) = PI * hh * hj
            wc2(4) = area_c2(4) * plt_body2 * ro_steel      '[kg] weight 3P pipe

            'weight 3P cone
            hh = _cyl1_dim(14) * _db2                       '[m] Length 3P cone
            hj = _cyl1_dim(12) * _db2                       '[m] grote diameter 3P pijp
            hk = _cyl1_dim(15) * _db2                       '[m] kleine diameter 3P pijp
            area_c2(5) = PI * (hj + hk) / 2 * hh
            wc2(5) = area_c2(5) * plt_body2 * ro_steel      '[kg] weight 3P pipe

            For i = 0 To wc2.Length - 1
                c_weight2 += wc2(i)                          'Total weight
                c_area2 += area_c2(i)
            Next

            c_weight2 *= 1.1                                '10% weight flanges +safety

            total_instal = c_weight1 * NumericUpDown20.Value  'parallel stage #1
            total_instal += c_weight2 * NumericUpDown33.Value 'parallel stage #2

            sheet_metal_wht = total_instal * 1.45                   'Gross Sheet metal weight
            p304 = sheet_metal_wht * NumericUpDown45.Value    'rvs 304
            p316 = sheet_metal_wht * NumericUpDown44.Value    'rvs 316

            '-------- present -----------

            TextBox99.Text = total_instal.ToString("F0")
            TextBox83.Text = sheet_metal_wht.ToString("F0")
            TextBox114.Text = p304.ToString("F0")
            TextBox115.Text = p316.ToString("F0")
            TextBox61.Text = c_weight1.ToString("F0")            'Total weight
            TextBox72.Text = c_weight2.ToString("F0")

            '------ area's cyclone 1----
            TextBox195.Text = area_c1(0).ToString("F2")          'Top Area [m2]
            TextBox196.Text = area_c1(1).ToString("F2")          'Body Area [m2]
            TextBox197.Text = area_c1(2).ToString("F2")          'Cone Area [m2]
            TextBox199.Text = area_c1(3).ToString("F2")          'Discharge pipe Area [m2]
            TextBox201.Text = area_c1(4).ToString("F2")          '3P pipe Area [m2]
            TextBox203.Text = area_c1(5).ToString("F2")          '3P cone Area [m2]
            TextBox135.Text = c_area1.ToString("F2")             'Area [m2]

            '------ area's cyclone 2----
            TextBox194.Text = area_c2(0).ToString("F2")          'Top Area [m2]
            TextBox193.Text = area_c2(1).ToString("F2")          'Body Area [m2]
            TextBox192.Text = area_c2(2).ToString("F2")          'Cone Area [m2]
            TextBox198.Text = area_c2(3).ToString("F2")          'Discharge pipe Area [m2]
            TextBox200.Text = area_c2(4).ToString("F2")          '3P pipe Area [m2]
            TextBox202.Text = area_c2(5).ToString("F2")          '3P cone Area [m2]
            TextBox127.Text = c_area2.ToString("F2")             'Area [m2]


        End If
    End Sub

    Private Sub Calc_Tang_outlet()
        Dim words As String()
        Dim _db1, _db2 As Double                    '[mm] outlet cyclone = inlet tangential
        Dim tan1(11) As Double                      '[mm] 
        Dim tan2(11) As Double
        Dim factor1(9) As Double                    '[-]
        Dim factor2(9) As Double                    '[-]

        If numericUpDown5.Value > 1 AndAlso NumericUpDown3.Value > 1 Then
            _db1 = CDbl(numericUpDown5.Value) * 1000          '[mm] dia cyclone stage 1
            _db2 = CDbl(NumericUpDown34.Value) * 1000         '[mm] dia cyclone stage 2

            '-------- dimension cyclone stage #1
            If ComboBox1.SelectedIndex < Tangent_out_dimensions.Length - 1 Then
                words = Tangent_out_dimensions(ComboBox1.SelectedIndex).Split(CType(";", Char()))
                For hh = 1 To factor1.Length - 2
                    factor1(hh) = CDbl(words(hh))          'Tangent outlet dimensions
                Next
            Else
                For hh = 1 To factor1.Length - 2
                    factor1(hh) = 0          'Tangent outlet dimensions
                Next
            End If


            '-------- dimension cyclone stage #2
            If ComboBox2.SelectedIndex < Tangent_out_dimensions.Length - 1 Then
                words = Tangent_out_dimensions(ComboBox2.SelectedIndex).Split(CType(";", Char()))
                For hh = 1 To factor2.Length - 2
                    factor2(hh) = CDbl(words(hh))          'Tangent outlet dimensions
                Next
            Else
                For hh = 1 To factor2.Length - 2
                    factor2(hh) = 0          'Tangent outlet dimensions
                Next
            End If

            '--------- 1st stage cyclone -----
            tan1(1) = _db1 * factor1(1)         'Diameter
            tan1(2) = _db1 * factor1(2)         'uitlaat breedte
            tan1(3) = _db1 * factor1(3)         'Uitlaat hoogte
            tan1(4) = _db1 * factor1(4)         'uitlaat lengte
            tan1(5) = _db1 * factor1(5)         'uitlaat hartmaat
            tan1(6) = _db1 * factor1(6)         'Steekmaat radii
            tan1(7) = _db1 * factor1(7)         'Radius 1
            tan1(8) = _db1 * factor1(8)         'Radius gooze neck
            tan1(9) = tan1(7) - 1 * tan1(6)     'Radius 2
            tan1(10) = tan1(7) - 2 * tan1(6)    'Radius 3

            '--------- 2nd stage cyclone -----
            tan2(1) = _db2 * factor2(1)         'Diameter
            tan2(2) = _db2 * factor2(2)         'uitlaat breedte
            tan2(3) = _db2 * factor2(3)         'Uitlaat hoogte
            tan2(4) = _db2 * factor2(4)         'uitlaat lengte
            tan2(5) = _db2 * factor2(5)         'uitlaat hartmaat
            tan2(6) = _db2 * factor2(6)         'Steekmaat radii
            tan2(7) = _db2 * factor2(7)         'Radius 1
            tan2(8) = _db2 * factor2(8)         'radius gooze neck
            tan2(9) = tan2(7) - 1 * tan1(6)     'Radius 2
            tan2(10) = tan2(7) - 2 * tan1(6)    'Radius 3

            '--------- 1st stage cyclone -----
            TextBox81.Text = CType(ComboBox1.SelectedItem, String)      'Cycloon type        
            TextBox82.Text = tan1(1).ToString("F0")         'Diameter
            TextBox154.Text = tan1(2).ToString("F0")        'uitlaat breedte
            TextBox155.Text = tan1(3).ToString("F0")        'Uitlaat hoogte
            TextBox156.Text = tan1(4).ToString("F0")        'uitlaat lengte
            TextBox157.Text = tan1(5).ToString("F0")        'uitlaat hartmaat
            TextBox158.Text = tan1(6).ToString("F0")        'Steekmaat radii
            TextBox159.Text = tan1(7).ToString("F0")        'Radius 1
            TextBox160.Text = tan1(8).ToString("F0")        'Radius gooze neck
            TextBox172.Text = tan1(9).ToString("F0")        'Radius 2
            TextBox173.Text = tan1(10).ToString("F0")       'Radius 3

            '--------- 2nd stage cyclone -----
            TextBox161.Text = CType(ComboBox2.SelectedItem, String)      'Cycloon type
            TextBox162.Text = tan2(1).ToString("F0")        'Diameter
            TextBox163.Text = tan2(2).ToString("F0")        'uitlaat breedte
            TextBox164.Text = tan2(3).ToString("F0")        'Uitlaat hoogte
            TextBox165.Text = tan2(4).ToString("F0")        'uitlaat lengte
            TextBox166.Text = tan2(5).ToString("F0")        'uitlaat hartmaat
            TextBox167.Text = tan2(6).ToString("F0")        'Steekmaat radii
            TextBox168.Text = tan2(7).ToString("F0")        'Radius 1
            TextBox169.Text = tan2(8).ToString("F0")        'Radius gooze neck
            TextBox170.Text = tan2(9).ToString("F0")        'Radius 2
            TextBox171.Text = tan2(10).ToString("F0")       'Tadius 3
        End If
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        '--- Save case ----
        Save_present_case_to_array()
    End Sub
    Private Sub Save_present_case_to_array()
        Dim ks As Integer
        'Save data of screen into the _cees array

        ks = CInt(NumericUpDown30.Value)       'Case number
        Fill_array_from_screen(ks)
    End Sub

    Private Sub NumericUpDown30_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown30.ValueChanged
        Dim ks As Integer

        '==== Different case is selected =======
        ProgressBar1.Visible = True         'Show the progress bar
        ProgressBar1.Value = 60
        GroupBox9.Update()                  'Show the progress bar

        ks = CInt(NumericUpDown30.Value)    'Case number
        Update_Screen_from_array(ks)        'Case number
        _cees(0) = _cees(ks)                'For calculation

        ProgressBar1.Visible = False
    End Sub

    Private Sub Update_Screen_from_array(zz As Integer)
        Dim p1_rel As Double
    
        If init Then
            SuspendLayout()
            '----------- General (not calculated) data------------------
            TextBox28.Text = _cees(1).Quote_no                 'Quote number (not case dependent)
            TextBox29.Text = _cees(1).Tag_no                   'The Tag number (not case dependent)
            TextBox53.Text = _cees(zz).case_name               'Case name

            Chck_value(NumericUpDown1, CDec(_cees(zz).FlowT))       'Air flow total
            Chck_value(NumericUpDown4, CDec(_cees(zz).dust1_Am3))   'Dust inlet [g/Am3] 

            ComboBox1.SelectedIndex = _cees(zz).Ct1                 'Cyclone type stage #1
            ComboBox2.SelectedIndex = _cees(zz).Ct2                 'Cyclone type stage #2

            'If ComboBox1.Items.Count > 0 Then ComboBox1.SelectedIndex = 0
            'If ComboBox2.Items.Count > 0 Then ComboBox2.SelectedIndex = 0

            Chck_value(NumericUpDown30, zz)                    'Case number
            Chck_value(NumericUpDown20, CDec(_cees(zz).Noc1))       'Cyclone in parallel #1
            Chck_value(NumericUpDown33, CDec(_cees(zz).Noc2))       'Cyclone in parallel #2
            Chck_value(numericUpDown5, CDec(_cees(zz).db1))         '[m] Diameter cyclone body #1
            Chck_value(NumericUpDown34, CDec(_cees(zz).db2))        '[m] Diameter cyclone body #2

            'Debug.WriteLine("Case zz= " & zz.ToString & ",  _cees(zz).Noc1= " & _cees(zz).Noc1.ToString)
            'Debug.WriteLine("Case zz= " & zz.ToString & ",  _cees(zz).Noc2= " & _cees(zz).Noc2.ToString)
            'Debug.WriteLine("Case zz= " & zz.ToString & ",  _cees(zz).db1= " & _cees(zz).db1.ToString)
            'Debug.WriteLine("Case zz= " & zz.ToString & ",  _cees(zz).db2= " & _cees(zz).db2.ToString)

            TextBox190.Text = _cees(zz).ro_gas.ToString             '[kg/m3] Density 
            Chck_value(numericUpDown2, CDec(_cees(zz).ro_solid))     '[kg/m3] Density 
            Chck_value(NumericUpDown18, CDec(_cees(zz).Temp))        '[c] Temperature 
            p1_rel = (_cees(zz).p1_abs - 101325) / 100.0             '[mbar]
            Chck_value(NumericUpDown19, CDec(p1_rel))                '[Pa abs]-->[mbar g] Pressure


            'Clear_dgv6()
            Fill_DGV6_from_input_array()
            Dump_log_to_box24()

            ResumeLayout()
        End If
    End Sub
    Private Sub Fill_DGV6_from_input_array()
        '[mu] Class upper particle diameter limit diameter
        '[%] Percentage van de inlaat stof belasting

        For row = 0 To DataGridView6.Rows.Count - 1
            DataGridView6.Rows(row).Cells(0).Value = _input(row).dia_big
            DataGridView6.Rows(row).Cells(1).Value = _input(row).class_load * 100
        Next
        DataGridView6.Refresh()
    End Sub

    Private Sub Dump_log_to_box24()
        Dim sb As New StringBuilder()

        TextBox24.Text = "----------------------------------------" & vbCrLf
        For zz = 0 To 4
            sb.AppendFormat("zz= {0} _cees(zz).Noc1= {1}", zz, _cees(zz).Noc1 & vbCrLf)
            sb.AppendFormat("zz= {0} _cees(zz).Noc2= {1}", zz, _cees(zz).Noc2 & vbCrLf)
            sb.AppendFormat("zz= {0} _cees(zz).db1= {1}", zz, _cees(zz).db1 & vbCrLf)
            sb.AppendFormat("zz= {0} _cees(zz).db2= {1}", zz, _cees(zz).db2 & vbCrLf & vbCrLf)
        Next

        For zz = 0 To 4
            sb.AppendFormat("= _input({1}).dia_big= ", zz, _input(zz).dia_big & vbCrLf)
        Next
        TextBox24.Text &= sb.ToString
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

    Private Sub Calc_stage1_2_comb(ks As Integer)

        Dim li(11) As Double
        Dim row As Integer
        Dim w19 As Double   'Dust load [kg/Nm3] 1stage
        Dim w20 As Double   'Dust load [kg/Nm3] 2stage

        If _cees(ks).stage1(row).d_ave > 0 Then   'For fast startup

            DataGridView4.ColumnCount = 10
            DataGridView4.Rows.Clear()
            DataGridView4.Rows.Add(111)
            DataGridView4.EnableHeadersVisualStyles = False                         'For backcolor
            DataGridView4.RowHeadersVisible = False

            DataGridView4.Columns(0).HeaderText = "Dia Class [mu]"                  'Chart
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
            li(2) = 100.0 * li(1) / _cees(ks).dust1_Nm3             'In psd diff [%]
            li(3) = 100.0 - li(2)                                 'In psd diff cumm[%]" S76 
            li(4) = _cees(ks).stage2(row).psd_dif               'In psd cum [%]
            li(5) = 100.0 - _cees(ks).stage2(row).psd_dif         'Loss1 pdscum [%]
            li(6) = _cees(ks).stage2(row).loss_abs_C * 10 * w20 'Loss abs [g/Nm3]
            li(7) = 1000.0 * li(6) / _cees(ks).sum_loss_C2        'Loss2 pds dif [%]
            li(8) = 100.0 - li(7)                                 'Loss2 pds.cum [%]
            li(9) = 100.0 * (li(1) - li(6)) / li(1)               'Eff stage1&2 [%]

            For col = 0 To 9  'Fill the DataGrid
                If Double.IsNaN(li(col)) Then li(col) = 0           'prevent silly results
                If li(col) < 0 OrElse li(col) > 10 ^ 5 Then li(col) = 0   'prevent silly results
                DataGridView4.Rows(row).Cells(col).Value = li(col).ToString("F5")
            Next

            '========- rest of the lines ========
            Dim qq As Double
            For row = 1 To DataGridView4.Rows.Count - 1                 'Fill the DataGrid
                li(0) = _cees(ks).stage1(row).d_ave                     'Dia aver [mu]
                li(1) = _cees(ks).stage1(row).psd_dif * 10.0 * w19        'In abs [g/Am3]" T76 * W19
                li(2) = 100.0 * li(1) / _cees(ks).dust1_Nm3               'In psd diff [%]
                li(3) = _cees(ks).stage1(row - 1).psd_cum_pro - li(2)   'In psd diff cumm[%]
                li(4) = _cees(ks).stage2(row).psd_dif                   'In psd cum [%]

                qq = CDbl(DataGridView4.Rows(row - 1).Cells(5).Value)
                li(5) = qq - _cees(ks).stage2(row).psd_dif              'Loss1 pdscum [%] 
                li(6) = _cees(ks).stage2(row).loss_abs_C * 10 * w20     'Loss abs [g/Am3]
                li(7) = 1000.0 * li(6) / _cees(ks).sum_loss_C2            'Loss2 pds dif [%]

                qq = CDbl(DataGridView4.Rows(row - 1).Cells(8).Value)
                li(8) = qq - li(7)                                      'Loss2 pds.cum [%] 
                li(9) = 100.0 * ((li(1) - li(6)) / li(1))                 'Eff stage1&2 [%]

                If li(8) = 0.0 Then li(9) = 100.0       'Loss is zero eff must be 100%

                '========== prevent silly results ======
                For col = 0 To 9  'Fill the DataGridview
                    If Double.IsNaN(li(col)) Then li(col) = 0               'prevent silly results
                    If li(col) < 0 OrElse li(col) > 10 ^ 5 Then li(col) = 0     'prevent silly results
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
        If loss > 100.0 OrElse loss < 0.0 Then MsgBox("Problem in line Vlookup_db")

        loss = 100.0 - loss
        For row = 1 To DataGridView4.Rows.Count - 1
            If loss <= CDbl(DataGridView4.Rows(row).Cells(9).Value) Then
                Return CDbl(DataGridView4.Rows(row).Cells(0).Value)
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
        If p1 < 100.0 Then p1 = 100.0             'Prevent devide by zero
        ro_normal = ro1 * (101325 / p1) * ((t1 + 273.15) / 273.15)
        Return (ro_normal)
    End Function

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click, TabPage10.Enter
        Draw_chart3(Chart3)             'Dust load correction
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Form2.Show()    'Fopspeen
    End Sub
    Private Sub Build_clear_dgv6()
        'This is the user input mechanism for entering the PSD data
        With DataGridView6
            .ColumnCount = 2
            .Rows.Clear()
            .Rows.Add(no_PDS_inputs - 1)
            .EnableHeadersVisualStyles = False   'For backcolor
            .RowHeadersVisible = False

            .Columns(0).DefaultCellStyle.Format = "N3"
            .Columns(1).DefaultCellStyle.Format = "N3"

            .Columns(0).HeaderText = "Upper Dia [um]"
            .Columns(1).HeaderText = "Cumm [%] tot wght"

            .Columns(0).Width = 95
            .Columns(1).Width = 95


            For Each row As DataGridViewRow In .Rows
                row.Cells(0).Value = 0
                row.Cells(1).Value = 0
            Next
        End With

    End Sub

    Private Sub DataGridView6_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellEndEdit
        'New data entered by user
        Calc_sequence()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        PSD_maltodesxtrine() 'Chineese determined PSD
        Calc_sequence()
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        PSD_maltodextrine_2()
        Calc_sequence()
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        PSD_Whey()
        Calc_sequence()
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        PSD_DSM_polymere()
        Calc_sequence()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        AA850_psd()
        Calc_sequence()
    End Sub
    Private Sub AA850_psd()
        TextBox28.Text = "AA850 test"
        TextBox29.Text = "--"
        TextBox53.Text = "--"
        NumericUpDown1.Value = 55000      '[Am3/h] Flow
        NumericUpDown18.Value = 60        '[c]
        NumericUpDown19.Value = -80       '[mbar] 
        numericUpDown2.Value = 1500       '[kg/m3] density
        TextBox190.Text = "0.903"               '[kg/m3] ro air
        NumericUpDown4.Value = 139        '[g/Am3]
        NumericUpDown30.Value = 1         '[-] Case number

        NumericUpDown20.Value = 80        '[-] parallel cycloon
        ComboBox1.SelectedIndex = 9       'AA850 stage #1
        numericUpDown5.Value = CDec(0.3)        '[m] diameter cycloon

        NumericUpDown33.Value = 80         '[-] parallel cycloon
        ComboBox2.SelectedIndex = 9        'AA850 stage #2
        NumericUpDown34.Value = CDec(0.3)        '[m] diameter cycloon

        '======== Fill the DVG with PSD example data =======
        Fill_dgv6_example(AA_excel)
    End Sub

    Private Sub PSD_maltodesxtrine()
        TextBox28.Text = "Cargill China"
        TextBox29.Text = "Q20.1021"
        TextBox53.Text = "--"
        NumericUpDown1.Value = 75000        '[Am3/h] Flow
        NumericUpDown18.Value = 120         '[c]
        NumericUpDown19.Value = -30         '[mbar] 
        numericUpDown2.Value = 1200         '[kg/m3] density
        TextBox190.Text = "0.8977"                '[kg/m3] ro air
        NumericUpDown4.Value = 20          '[g/Am3]
        NumericUpDown30.Value = 1           '[-] Case number

        NumericUpDown20.Value = 120        '[-] parallel cycloon
        ComboBox1.SelectedIndex = 9         'AA850 stage #1
        numericUpDown5.Value = CDec(0.3)    '[m] diameter cycloon

        NumericUpDown33.Value = 120         '[-] parallel cycloon
        ComboBox2.SelectedIndex = 9         'AA850 stage #2
        NumericUpDown34.Value = CDec(0.3)   '[m] diameter cycloon

        '======== Fill the DVG with PSD example data =======
        Fill_dgv6_example(maltodextrine_psd)
    End Sub

    Private Sub PSD_maltodextrine_2()
        TextBox28.Text = "Cargill China"
        TextBox29.Text = "Q20.1021"
        TextBox53.Text = "--"
        NumericUpDown1.Value = 96600       '[Am3/h] Flow
        NumericUpDown18.Value = 118         '[c]
        NumericUpDown19.Value = -3          '[mbar] 
        numericUpDown2.Value = 800          '[kg/m3] density
        TextBox190.Text = "0.859"               '[kg/m3] ro air
        NumericUpDown4.Value = CDec(9.35)   '[g/Am3]
        NumericUpDown30.Value = 1          '[-] Case number

        NumericUpDown20.Value = 2          '[-] parallel cycloon
        ComboBox1.SelectedIndex = 5         'AC850 stage #1
        numericUpDown5.Value = CDec(2.8)    '[m] diameter cycloon

        NumericUpDown33.Value = 144         '[-] parallel cycloon
        ComboBox2.SelectedIndex = 9         'AA850 stage #2
        NumericUpDown34.Value = CDec(0.3)   '[m] diameter cycloon

        '======== Fill the DVG with PSD example data =======
        Fill_dgv6_example(maltodextrine_psd_suresh)
    End Sub

    Private Sub PSD_Whey()
        TextBox28.Text = "PSD Whey"
        TextBox29.Text = "Q20.1018"
        TextBox53.Text = "--"
        NumericUpDown1.Value = 155952      '[Am3/h] Flow
        NumericUpDown18.Value = 77          '[c]
        NumericUpDown19.Value = -20         '[mbar] 
        numericUpDown2.Value = 1600         '[kg/m3] density
        TextBox190.Text = "1.278"               '[kg/m3] ro air
        NumericUpDown4.Value = 20           '[g/Am3]
        NumericUpDown30.Value = 1           '[-] Case number

        NumericUpDown20.Value = 3           '[-] parallel cycloon
        ComboBox1.SelectedIndex = 2         'AC435 stage #1
        numericUpDown5.Value = CDec(2.2)    '[m] diameter cycloon

        NumericUpDown33.Value = 6          '[-] parallel cycloon
        ComboBox2.SelectedIndex = 5         'AC850 stage #2
        NumericUpDown34.Value = CDec(2.2)   '[m] diameter cycloon

        '======== Fill the DVG with PSD example data =======
        Fill_dgv6_example(psd_whey_A6605)

    End Sub
    Private Sub PSD_typical_Corn()
        '======== Fill the DVG with Corn example data (difficult stuff) =======
        Fill_dgv6_example(psd_corn)
    End Sub

    Private Sub PSD_typical_Chickpea_starch()
        '======== Fill the DVG with Chickpea_starch =======
        Fill_dgv6_example(psd_chickpea_starch)
    End Sub

    Private Sub PSD_DSM_polymere()
        '======== Fill the DVG with DSM polymere example data =======
        Fill_dgv6_example(DSM_psd_example)
    End Sub
    Private Sub Fill_dgv6_example(ww As String())
        Dim words As String()
        With DataGridView6
            For row = 0 To .Rows.Count - 1
                If row < ww.Length Then
                    words = ww(row).Split(CType(";", Char()))
                    .Rows(row).Cells(0).Value = words(0).Trim
                    .Rows(row).Cells(1).Value = words(1).Trim
                Else
                    .Rows(row).Cells(0).Value = 0
                    .Rows(row).Cells(1).Value = 0
                End If
            Next
        End With
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        PSD_GvG_excel()
        Calc_sequence()
    End Sub

    Private Sub PSD_GvG_excel()
        TextBox28.Text = "GvG_test_excel"
        TextBox29.Text = "--"
        TextBox53.Text = "--"
        NumericUpDown1.Value = 58900        '[Am3/h]
        NumericUpDown18.Value = 45          '[c]
        NumericUpDown19.Value = -30         '[mbar]
        numericUpDown2.Value = 1500         '[kg/m3]
        TextBox190.Text = "1.073"               '[kg/m3]
        NumericUpDown4.Value = 77           '[g/Am3]
        NumericUpDown30.Value = 1           '[-] Case number
        NumericUpDown20.Value = 3           '[-] parallel cycloon
        ComboBox1.SelectedIndex = 2         'AC435 stage #1
        numericUpDown5.Value = CDec(1.25)   '[m] diameter cycloon
        NumericUpDown33.Value = 6           '[-] parallel cycloon
        ComboBox2.SelectedIndex = 5        'AC850 stage #2
        NumericUpDown34.Value = CDec(1.25)  '[mm] diameter cycloon

        '======== Fill the DVG6 with DSM polymere example data =======
        Fill_dgv6_example(GvG_excel)
        'Debug.WriteLine("PSD_GvG_excel()")
    End Sub

    Private Sub Screen_contrast()
        '====This fuction is to increase the readability=====
        '==== of the red text ===============================
        Dim all_txt, all_num, all_lab As New List(Of Control)

        '-------- find all Text box controls -----------------
        FindControlRecursive(all_txt, Me, GetType(TextBox))   'Find the control
        For i = 0 To all_txt.Count - 1
            Dim grbx As TextBox = CType(all_txt(i), TextBox)
            If grbx.BackColor.Equals(Color.Red) Then
                grbx.Enabled = True
                grbx.ForeColor = Color.White
            Else
                grbx.ForeColor = Color.Black
            End If
        Next

        '-------- find all numeric controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        For i = 0 To all_num.Count - 1
            Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
            If grbx.BackColor.Equals(Color.Red) Then
                grbx.ForeColor = Color.White
            Else
                grbx.ForeColor = Color.Black
            End If
        Next

        '-------- find all label controls -----------------
        FindControlRecursive(all_lab, Me, GetType(Label))   'Find the control
        For i = 0 To all_lab.Count - 1
            Dim grbx As Label = CType(all_lab(i), Label)
            If grbx.BackColor.Equals(Color.Red) Then
                grbx.Enabled = True
                grbx.ForeColor = Color.White
            Else
                grbx.ForeColor = Color.Black
            End If
        Next
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

    Private Sub Rights_Control()
        Dim id As String        'This is the present user

        id = Trim(Environment.UserName)     'User name 
        id = LCase(id)

        '========= Disable page Tabs for everybody ===============
        TabControl1.TabPages.Remove(TabPage8)       'Logging
        TabControl1.TabPages.Remove(TabPage10)      'High Dust load
        PictureBox4.Visible = False
        TextBox20.Visible = False
        TextBox174.Visible = False

        If (id = "gp" OrElse id = "gerritp" OrElse id = "user") Then
            TabControl1.TabPages.Add(TabPage8)       'Logging
            TabControl1.TabPages.Add(TabPage10)      'High Dust load
            PictureBox4.Visible = True
            TextBox20.Visible = True
            TextBox174.Visible = True
        End If
    End Sub

    Private Sub DataGridView6_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView6.KeyDown
        '=========== Paste action for DGV6 ================
        If e.Control AndAlso e.KeyCode = Keys.V Then
            Dim row As Integer = 0
            For Each line As String In Clipboard.GetText.Split(CChar(vbNewLine))
                If line.Trim.Length > 0 Then

                    Dim item As String() = line.Split(vbTab(0)).Select(Function(X) X.Trim).ToArray
                    item(0) = item(0).Replace(",", ".")
                    item(1) = item(1).Replace(",", ".")
                    DataGridView6.Rows(row).Cells(0).Value = item(0)
                    DataGridView6.Rows(row).Cells(1).Value = item(1)
                    row += 1
                End If
            Next

        End If
        Calc_sequence()
        Debug.WriteLine("DataGridView6_KeyDown done")
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Build_clear_dgv6()  'Input Clear the all grid cells
        Calc_sequence()
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        Form3.Size = New Size(920, 600)     '(Breed, Hoog)

        Form3.Text = "PSD Starch"
        Form3.PictureBox1.Image = My.Resources.PSD_Starch
        Form3.Show()
        Form3.TopMost = True
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        Form3.Size = New Size(920, 620)     '(Breed, Hoog)
        Form3.Text = "PSD Starch"
        Form3.PictureBox1.Image = My.Resources.PSD_various
        Form3.Show()
        Form3.TopMost = True
    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        Form3.Size = New Size(650, 800)     '(Breed, Hoog)
        Form3.Text = "Micrographs Starch"
        Form3.PictureBox1.Image = My.Resources.Micrographs
        Form3.Show()
        Form3.TopMost = True
    End Sub

    Private Sub PictureBox11_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click
        Form3.Size = New Size(650, 800)     '(Breed, Hoog)
        Form3.Text = "Chickpea PSD"
        Form3.PictureBox1.Image = My.Resources.Chickpea_psd
        Form3.Show()
        Form3.TopMost = True
    End Sub


    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        '======= Invert the Cumm weight colums =====
        For Each row As DataGridViewRow In DataGridView6.Rows
            If Not IsNothing(row.Cells(0).Value) AndAlso Not IsNothing(row.Cells(1).Value) Then
                If Not IsNothing(row.Cells(1).Value) Then row.Cells(1).Value = 100.0 - CDbl(row.Cells(1).Value)

                '===== diameter NULL than also weight is NUL ======
                If CDbl(row.Cells(0).Value) < 0.00001 Then
                    row.Cells(0).Value = 0
                    row.Cells(1).Value = 0
                End If
            End If
        Next
        Calc_sequence()
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        TextBox28.Text = "Q20.1140"
        TextBox29.Text = "PSD Potato,Balaji "
        TextBox53.Text = "Summer"
        NumericUpDown1.Value = 3300         '[Am3/h] Flow
        NumericUpDown18.Value = 51          '[c]
        NumericUpDown19.Value = -46         '[mbar] 
        numericUpDown2.Value = 1600         '[kg/m3] density
        TextBox190.Text = "1.013"           '[kg/m3] ro air
        'numericUpDown14.Value = CDec(0.019548) '[mPas=cP] visco air
        NumericUpDown4.Value = CDec(109.8)  '[g/Am3]
        NumericUpDown30.Value = 1           '[-] Case number

        NumericUpDown20.Value = 1           '[-] parallel cycloon
        ComboBox1.SelectedIndex = 2         'AC435 stage #1
        numericUpDown5.Value = CDec(0.45)   '[m] diameter cycloon

        NumericUpDown33.Value = 2           '[-] parallel cycloon
        ComboBox2.SelectedIndex = 5         'AC850 stage #2
        NumericUpDown34.Value = CDec(0.45)  '[m] diameter cycloon

        '======== Fill the DVG with PSD example data =======
        Fill_dgv6_example(psd_potato_flash_drier)
        Calc_sequence()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        PSD_typical_Corn()
        Calc_sequence()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        PSD_typical_Chickpea_starch()
        Calc_sequence()
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click, TabPage13.Enter
        Draw_chart4(Chart4)             'PSD selected product
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click, NumericUpDown8.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown13.ValueChanged, ComboBox3.SelectedIndexChanged, TabPage15.Enter, RadioButton9.CheckedChanged, RadioButton8.CheckedChanged, RadioButton7.CheckedChanged
        If TextBox37.Text.Length > 1 Then   'prevent startup exceptions
            Design_stress()
            Sync_input1()
            Calc_cyl_shell742()         'Cylindrical shell
            Calc_conical_shell764()     'Conus shell
            Calc_Junction766()          'Junction large end
        End If
    End Sub
    Private Sub Sync_input1()
        Dim dia, apex As Decimal
        Dim Lcyl, Lcone As Decimal
        Dim dia_discharge As Decimal


        Label291.Text = "No sync, free selection"
        Select Case True
            Case RadioButton8.Checked
                Label291.Text = "Sync cyclone stage 1"
                Decimal.TryParse(TextBox37.Text, dia)
                Decimal.TryParse(TextBox150.Text, apex)
                Decimal.TryParse(TextBox10.Text, Lcyl)
                Decimal.TryParse(TextBox11.Text, Lcone)
                Decimal.TryParse(TextBox6.Text, dia_discharge)

                Lcone = CDec(dia * 0.5 / Tan(apex / 180 * PI))
            Case RadioButton9.Checked
                Label291.Text = "Sync cyclone stage 2"
                Decimal.TryParse(TextBox74.Text, dia)
                Decimal.TryParse(TextBox151.Text, apex)
                Decimal.TryParse(TextBox93.Text, Lcyl)
                Decimal.TryParse(TextBox94.Text, Lcone)
                Decimal.TryParse(TextBox89.Text, dia_discharge)
                Lcone = CDec(dia * 0.5 / Tan(apex / 180 * PI))
        End Select

        If RadioButton8.Checked OrElse RadioButton9.Checked Then
            ComboBox3.SelectedIndex = 0             'Weld factor
            Set_numeric_control_value(NumericUpDown10, dia * 1000)              'Shell OD
            Set_numeric_control_value(NumericUpDown15, dia * 1000)              'Shell OD
            Set_numeric_control_value(NumericUpDown13, apex)                    '1/2 apex
            Set_numeric_control_value(NumericUpDown57, apex)                    '1/2 apex
            Set_numeric_control_value(NumericUpDown9, 200)                      'Height klöpper bodem
            Set_numeric_control_value(NumericUpDown23, Lcyl * 1000)             '[mm] Cylinder
            Set_numeric_control_value(NumericUpDown7, Lcone * 1000)             '[mm] Cone
            Set_numeric_control_value(NumericUpDown16, dia_discharge * 1000)    '[mm] Cone
            Set_numeric_control_value(NumericUpDown11, NumericUpDown18.Value)   '[c] Temp

            '---- change color ---
            ComboBox3.BackColor = Color.White           'Weld factor
            NumericUpDown10.BackColor = Color.White     'Shell OD
            NumericUpDown15.BackColor = Color.White     'Shell OD
            NumericUpDown13.BackColor = Color.White     '1/2 apex
            NumericUpDown57.BackColor = Color.White     '1/2 apex
            NumericUpDown9.BackColor = Color.White      'Height klöpper bodem
            NumericUpDown23.BackColor = Color.White     '[mm] Cylinder
            NumericUpDown7.BackColor = Color.White      '[mm] Cone
            NumericUpDown16.BackColor = Color.White     '[mm] Dia oulet pipe
            NumericUpDown11.BackColor = Color.White     '[c] Temperature
        Else
            ComboBox3.BackColor = Color.Yellow          'Weld factor
            NumericUpDown10.BackColor = Color.Yellow    'Shell OD
            NumericUpDown15.BackColor = Color.Yellow    'Shell OD
            NumericUpDown13.BackColor = Color.Yellow    '1/2 apex
            NumericUpDown57.BackColor = Color.Yellow    '1/2 apex
            NumericUpDown9.BackColor = Color.Yellow     'Height klöpper bodem
            NumericUpDown23.BackColor = Color.Yellow    '[mm] Cylinder
            NumericUpDown7.BackColor = Color.Yellow     '[mm] Cone
            NumericUpDown16.BackColor = Color.Yellow    '[mm] Dia oulet pipe
            NumericUpDown11.BackColor = Color.Yellow     '[c] Temperature
        End If
    End Sub

    '7.4.2 Cylindrical shells internal pressure
    Private Sub Calc_cyl_shell742()
        Dim De, Di, Dm, ea, z_joint, e_wall, Pmax, valid_check As Double

        If (ComboBox3.SelectedIndex > -1) Then          'Prevent exceptions
            Double.TryParse(joint_eff(ComboBox3.SelectedIndex), z_joint)      'Joint efficiency
        End If

        De = NumericUpDown15.Value  'OD [mm]
        ea = NumericUpDown8.Value  'Wall thicknes [mm]
        Di = De - 2 * ea            'ID [mm]

        Dm = (De + Di) / 2                          'Average Dia
        Pmax = 2 * _fs * z_joint * ea / Dm          'Max pressure equation 7.4.3 

        e_wall = _P * De / (2 * _fs * z_joint + _P) 'equation 7.4.2 Required wall thickness
        valid_check = Round(e_wall / De, 4)

        '--------- present results--------
        TextBox211.Text = Di.ToString("F1")           '[mm] ID cone
        TextBox210.Text = Round(e_wall, 4).ToString   'required wall [mm]
        TextBox209.Text = valid_check.ToString
        TextBox208.Text = _P.ToString("F2")           '[MPa]
        TextBox206.Text = (_P * 10).ToString("F2")    '[Bar]
        TextBox207.Text = _fs.ToString
        TextBox212.Text = (Pmax * 10).ToString("F2")  '[Bar]

        '---------- Check-----
        TextBox209.BackColor = CType(IIf(valid_check > 0.16, Color.Red, Color.LightGreen), Color)
        TextBox212.BackColor = CType(IIf(Pmax < _P, Color.Red, Color.LightGreen), Color)
    End Sub
    '7.6.4 Conical shells 
    Private Sub Calc_conical_shell764()
        Dim α As Double 'Half apex cone
        Dim z_joint As Double
        Dim De, Di, ea, e_con, e_cone As Double
        Dim pmaxx As Double 'max pressure
        Dim Dm As Double


        If (ComboBox3.SelectedIndex > -1) Then          'Prevent exceptions
            Double.TryParse(joint_eff(ComboBox3.SelectedIndex), z_joint)      'Joint efficiency

            De = NumericUpDown15.Value                  'OD
            ea = NumericUpDown8.Value                   'Wall thicknes
            Di = De - 2 * ea            'ID
            α = NumericUpDown13.Value / 180 * PI        'Half apex in radials
            Dm = (De + Di) / 2                          'Average diameter
            e_cone = NumericUpDown42.Value              'Cone wall

            '----------- cone wall thickness ----------
            e_con = _P * Di / (2 * _fs * z_joint - _P)  'equation (7.6-2) Required wall thickness
            e_con *= 1 / Cos(α)

            '---------- max pressure ---------------
            pmaxx = 2 * _fs * z_joint * e_cone * Cos(α) / Dm   'Max pressure equation (7.6-4) 

            '--------- present results--------
            TextBox205.Text = Round(e_con, 2).ToString          'required cone wall [mm]
            TextBox204.Text = (pmaxx * 10).ToString("F2")       '[MPa]-->[Bar]
            TextBox211.Text = Di.ToString("F1")                 '[mm] ID cone big end
            TextBox213.Text = ea.ToString("F1")                 '[mm] wall

            '---------- Check-----
            TextBox204.BackColor = CType(IIf(pmaxx < _P, Color.Red, Color.LightGreen), Color)
        End If
    End Sub

    '7.6.6 Junction between the large end of a cone and a cylinder without a knuckle
    Private Sub Calc_Junction766()
        Dim α As Double     'is Half apex cone
        Dim β As Double     'is a factor defined in 7.6.6;
        Dim Dc As Double    'diameter large end cone
        Dim ej As Double    'is a required or analysis thickness at a junction at the large end of a cone
        Dim ej1 As Double

        Dc = NumericUpDown15.Value              '[mm] OD cone large end
        α = NumericUpDown13.Value / 180 * PI    '[-] Half apex in radials

        ej = 40             '[mm] Initial thickness, Now iterate

        For i = 1 To 1000

            '----------- factor β ---------------------
            β = 1 / 3 * Sqrt(Dc / ej)                '(7.6-11)
            β *= Tan(α) / (1 + 1 / Sqrt(Cos(α)))
            β -= 0.15

            '----------- factor ej ---------------------
            ej1 = _P * Dc * β / (2 * _fs)            '(7.6-12)

            If ej < ej1 Then
                ej *= 1.03
            Else
                ej *= 0.97
            End If
            If Abs(ej - ej1) < 0.01 Then
                i = 1000
                TextBox188.BackColor = Color.LightGreen
            Else
                TextBox188.BackColor = Color.Red
            End If
        Next

        '--------- present results--------
        TextBox236.Text = β.ToString("F2")     '[-] Factor '(7.6-11)
        TextBox188.Text = ej.ToString("F1")    '[mm] required cone wall 
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click, NumericUpDown12.ValueChanged, NumericUpDown11.ValueChanged, ComboBox5.SelectedIndexChanged, ComboBox4.SelectedIndexChanged, RadioButton6.CheckedChanged, RadioButton5.CheckedChanged, RadioButton4.CheckedChanged, RadioButton3.CheckedChanged
        Design_stress()
    End Sub

    Private Sub Design_stress()
        Dim sf As Double = 1        'Safety factor init value
        Dim temperature As Double   'temperature
        Dim words As String()
        Dim y50, y100, y150, y200, y250, y300, y350, y400 As Double
        Dim ΔT As Double

        words = steel(ComboBox5.SelectedIndex + 1).Split(separators, StringSplitOptions.None)
        TextBox226.Text = words(1)
        TextBox225.Text = words(2)
        TextBox224.Text = words(3)
        TextBox223.Text = words(4)
        TextBox222.Text = words(5)
        TextBox221.Text = words(6)
        TextBox220.Text = words(7)
        TextBox219.Text = words(8)
        TextBox232.Text = words(13) 'cs or ss
        Double.TryParse(words(1), y50)
        Double.TryParse(words(2), y100)
        Double.TryParse(words(3), y150)
        Double.TryParse(words(4), y200)
        Double.TryParse(words(5), y250)
        Double.TryParse(words(6), y300)
        Double.TryParse(words(7), y350)
        Double.TryParse(words(8), y400)

        temperature = NumericUpDown11.Value        '[c]

        If (ComboBox5.SelectedIndex > -1) AndAlso (ComboBox4.SelectedIndex > -1) Then  'Prevent exceptions
            Select Case True
                Case 50 >= temperature
                    _f02 = CDec(y50)
                Case 100 >= temperature
                    ΔT = 50 - temperature
                    _f02 = Calc_design_stress(y50, y100, ΔT)
                Case 150 >= temperature
                    ΔT = 100 - temperature
                    _f02 = Calc_design_stress(y100, y150, ΔT)
                Case 200 >= temperature
                    ΔT = 150 - temperature
                    _f02 = Calc_design_stress(y150, y200, ΔT)
                Case 250 >= temperature
                    ΔT = 200 - temperature
                    _f02 = Calc_design_stress(y200, y250, ΔT)
                Case 300 >= temperature
                    ΔT = 250 - temperature
                    _f02 = Calc_design_stress(y250, y300, ΔT)
                Case 350 >= temperature
                    ΔT = 300 - temperature
                    _f02 = Calc_design_stress(y300, y350, ΔT)
                Case 400 >= temperature
                    ΔT = 350 - temperature
                    _f02 = Calc_design_stress(y350, y400, ΔT)
                Case temperature > 450
                    MessageBox.Show("Problem temperature too high")
            End Select

            _P = NumericUpDown12.Value                      'Calculation pressure [MPa=N/mm2]

            '------- chapter 6 rupture sensitivity ----
            ComboBox4.Enabled = False
            Select Case True
                Case String.Equals(TextBox232.Text, "cs")
                    ComboBox4.SelectedIndex = 0
                Case String.Equals(TextBox232.Text, "ss")
                    ComboBox4.SelectedIndex = 1
                Case Else
                    ComboBox4.SelectedIndex = 2
            End Select


            If String.Equals(TextBox232.Text, "cs") Then
                _Emod = (213.16 - 6.92 * temperature / 10 ^ 2 - 1.824 / 10 ^ 5 * temperature ^ 2) * 1000 '[N/mm2]
            Else
                _Emod = (201.66 - 8.48 * temperature / 10 ^ 2) * 1000      '[N/mm2]
            End If

            words = chap6(ComboBox4.SelectedIndex).Split(separators, StringSplitOptions.None)
            Double.TryParse(words(1), sf)               'Safety factor
            _fs = CDec(_f02 / sf)

            Select Case True
                Case RadioButton4.Checked
                    _fs *= 1        'PED article 3.3 (NO calc required)
                Case RadioButton5.Checked
                    _fs *= 1        'PED I,II,III
                Case RadioButton6.Checked
                    _fs *= 0.9      'PED IV
                Case RadioButton3.Checked
                    _fs = _f02      'EN 14460 6.2.1 (Shock resistant)
                    sf = 1.0
            End Select

            '-------- present -------------
            TextBox230.Text = (_P * 10 ^ 4).ToString        'Calculation pressure [mBar]
            TextBox233.Text = sf.ToString("F1")             'Safety factor
            TextBox231.Text = _f02.ToString("F0")           'Max allowed bend

            NumericUpDown14.Value = CDec(_fs)               '[N/mm2] Design stress
            TextBox133.Text = _f02.ToString("F0")           '[N/mm2] Yield stress
            TextBox229.Text = _Emod.ToString("F0")          '[N/mm2] Youngs modulus
            TextBox228.Text = _ν.ToString("F1")             'Poissons rate for steel
        End If
    End Sub
    Public Function Calc_design_stress(stress_A As Double, stress_B As Double, ΔT As Double) As Double
        Dim Δy As Double

        Δy = stress_B - stress_A
        Return (stress_A - (ΔT / 50 * Δy))
    End Function

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown57.ValueChanged, NumericUpDown48.ValueChanged, NumericUpDown23.ValueChanged, NumericUpDown10.ValueChanged, TabPage17.Enter, NumericUpDown46.ValueChanged
        For i = 0 To 2
            Design_stress()
            Calc_Cylinder_vacuum_852()
            'Calc_Round_plate()
            Calc_top_plate()
        Next
    End Sub
    Private Sub Calc_Cylinder_vacuum_852()
        'Vacumm cylindrical shell NO stiffeners
        Dim σe As Double
        Dim α As Double 'Half apex cone
        Dim De, Lcyl, h, Lcon As Double
        Dim matS As Double  'material safety
        Dim Tolerance As Double
        Dim Pr As Double    'calculated lower bound collapse pressure 
        Dim Py As Double    'pressure at which mean circumferential stress yields
        Dim Pm As Double    'theoretical elastic instability pressure for collapse of a perfect cylindrical
        Dim L As Double     'unsupported length of the shell
        Dim R_ As Double 'mean radius of a cylindrical or spherical shell
        Dim ε As Double     'mean elastic circumferential strain at collapse, see 8.5.2.2
        Dim ncyl As Double  'number of circumferential waves for an unstiffened part of a cylinder, see 8.5.2.2; 
        Dim Z As Double     'Formula (8.5.2-7) 
        Dim ea As Double    'shell wall thickness
        Dim x As Double     'Figure 8.5-5 — Values of Pt/PP versus Pm/PP 
        Dim PrPy As Double  'Figure 8.5-5 — Values of Pt/PP versus Pm/PP 

        '--- get data ----
        De = NumericUpDown10.Value      'OD shell
        Lcyl = NumericUpDown23.Value    'Cylinder length
        h = NumericUpDown9.Value       'Dished head height
        Lcon = NumericUpDown7.Value    'Cone length
        ea = NumericUpDown48.Value      'Shell wall thickness
        α = NumericUpDown57.Value       'Half apex cone
        R_ = De / 2                     'Radius shell

        '---- material --------
        matS = 1.5         'Safety factor (8.4.4-1) 

        '8.4.3 For shells made in austenitic steel, the nominal elastic limit shall be given by: 
        σe = _f02        'Mill certificate number
        If String.Equals(TextBox232.Text, "ss") Then
            σe /= 1.25
        Else
            σe /= 1.0
        End If


        '---- calculated lower bound collapse pressure obtained from Figure 8.5-5 ----------
        If α >= 30 Then
            L = Lcyl + 0.4 * h    '(8.5.2) Unsupported Length
        Else
            L = Lcyl + 0.4 * h + Lcon   '(8.5.3) Unsupported Length
        End If

        '--- pressure at which mean circumferential stress yields
        Py = σe * ea / R_    '(8.5.2-4) 

        Z = PI * R_ / L      '(8.5.2-7) 

        '---------- Find the smallest Pm ----------
        Dim Pm_small As Double = 9999

        For i = 2 To 20
            '--- calculate ε ----
            ε = Calc_ε(i, Z, R_, ea, _ν) '(8.5.2-6) 
            '--- theoretical elastic instability pressure for collapse of a perfect cylindrical
            Pm = _Emod * ea * ε / R_         '(8.5.2-5)
            If Pm < Pm_small Then
                Pm_small = Pm
                ncyl = i
            End If
        Next

        '--- now return to the smalles found case ----
        ε = Calc_ε(ncyl, Z, R_, ea, _ν) '(8.5.2-6) 
        '--- theoretical elastic instability pressure for collapse of a perfect cylindrical
        Pm = _Emod * ea * ε / R_         '(8.5.2-5)

        '---------------------------
        x = Pm / Py
        PrPy = -0.0016 * x ^ 4 + 0.031 * x ^ 3 - 0.2225 * x ^ 2 + 0.7227 * x - 0.0288

        Pr = PrPy * Py  'Calculated lower bound collapse pressure obtained from Figure 8.5-5

        '----------- Circularity tolerance  ----------
        Tolerance = 0.005 * Pr / (_P * matS) * 100  '[%](8.5.1-1) 

        '--------- present results--------
        TextBox244.Text = (_P * 10).ToString("F2")    '[MPa]-->[Bar]
        TextBox242.Text = σe.ToString("F1")           '[N/mm]
        TextBox218.Text = matS.ToString("F1")            '[-]
        TextBox243.Text = Tolerance.ToString("F2")    '[-]circular toleeance
        TextBox239.Text = (Py * 10).ToString("F2")    '[MPa]-->[Bar] at which mean circumferential stress yields
        TextBox238.Text = L.ToString("F0")            '[mm] unsupported length
        TextBox237.Text = (Pm * 10).ToString("F2")    '[MPa]-->[Bar]
        TextBox235.Text = Z.ToString("F1")            '[-]
        TextBox234.Text = ε.ToString("F5")            '[-]
        TextBox241.Text = _ν.ToString("F1")           '[-]
        TextBox240.Text = _Emod.ToString("F0")           '[-]
        TextBox217.Text = ncyl.ToString("F0")         '[-]
        TextBox216.Text = x.ToString("F1")            '[-]Pm/Py
        TextBox215.Text = PrPy.ToString("F2")         '[-]Pr/Py
        TextBox214.Text = (Pr * 10).ToString("F2")    '[MPa]-->[Bar]

        '---------- Check-----
        TextBox239.BackColor = CType(IIf(Py < _P, Color.Red, Color.LightGreen), Color)
        TextBox237.BackColor = CType(IIf(Pm < _P, Color.Red, Color.LightGreen), Color)
        TextBox214.BackColor = CType(IIf(Pr / matS < _P, Color.Red, Color.LightGreen), Color)
    End Sub
    Private Function Calc_ε(ncyl As Double, Z As Double, R_ As Double, ea As Double, ν As Double) As Double
        'Chapter External pressure 8.5
        Dim ε As Double
        ε = (ncyl ^ 2 - 1 + Z ^ 2) ^ 2      'Formula(8.5.2-6) 
        ε *= ea ^ 2 / (12 * R_ ^ 2 * (1 - ν ^ 2))
        ε += 1 / (ncyl ^ 2 / Z ^ 2 + 1) ^ 2
        ε *= 1 / (ncyl ^ 2 - 1 + Z ^ 2 / 2)
        Return (ε)
    End Function

    'Keep the numeric control within the Min and Max limits
    Private Sub Set_numeric_control_value(num As NumericUpDown, value As Decimal)
        Select Case True
            Case value <= num.Maximum AndAlso value >= num.Minimum
                num.Value = value
            Case value > num.Maximum
                num.Value = num.Maximum
            Case value < num.Minimum
                num.Value = num.Minimum
        End Select
    End Sub

    Private Sub Calc_top_plate()
        'Round with hole
        Dim dia, diahole As Double
        Dim a, b, t As Double
        Dim σm, ym As Double
        Dim x, k1, k2 As Double
        Dim wght As Double

        If NumericUpDown10.Value > 0 AndAlso (_P > 0) Then
            dia = NumericUpDown10.Value / 1000          '[m]
            diahole = NumericUpDown16.Value / 1000      '[m]
            t = NumericUpDown46.Value / 1000            '[m]
            a = dia / 2
            b = diahole / 2

            '============= determine k1, k2 =================
            x = a / b

            k1 = 0.0067 * x ^ 4 - 0.0584 * x ^ 3 + 0.0519 * x ^ 2 + 0.6132 * x - 0.4358
            k2 = 0.0127 * x ^ 4 - 0.131 * x ^ 3 + 0.3117 * x ^ 2 + 0.6069 * x - 0.219

            If x > 5 Then k1 = 0.815
            If x > 5 Then k2 = 2.2

            'MessageBox.Show("a= " & a.ToString & " b= " & b.ToString & " t=" & t.ToString & " e=" & ee.ToString)
            '------ bend stress ----
            σm = k2 * _P * a ^ 2 / (t ^ 2) '[N/mm2]

            '------ bend -------
            ym = k1 * _P * a ^ 4
            ym /= _Emod * t ^ 3
            ym *= 10 ^ 3                        '[mm]
            wght = PI * dia ^ 2 * t * 7850      '[kg]

            TextBox141.Text = x.ToString("F1")
            TextBox144.Text = σm.ToString("F0")
            TextBox140.Text = k1.ToString("F3")
            TextBox245.Text = k2.ToString("F3")
            TextBox246.Text = ym.ToString("F1")
            TextBox139.Text = wght.ToString("F0")

            '===== check ================
            TextBox144.BackColor = CType(IIf(σm > _fs, Color.Red, Color.LightGreen), Color)
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click, NumericUpDown24.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown17.Enter, CheckBox21.CheckedChanged, Chart5.Enter, CheckBox22.CheckedChanged
        Calc_plot_Distribution()
    End Sub
    Private Sub Calc_plot_Distribution()
        'https://numerics.mathdotnet.com/api/MathNet.Numerics.Distributions/index.htm
        'https://www.weibull.com/hotwire/issue47/relbasics47.htm

        Dim shape As Double = NumericUpDown17.Value
        Dim scale As Double = NumericUpDown24.Value

        Dim xx As Double
        If shape > 0 AndAlso scale > 0 Then
            With DataGridView5
                .ColumnCount = 2
                .Rows.Clear()
                .Rows.Add(300)
                .RowHeadersVisible = False
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                .Columns(0).HeaderText = "Dia [mu]"
                .Columns(1).HeaderText = "Bell [%]"
            End With
            With DataGridView7
                .ColumnCount = 2
                .Rows.Clear()
                .Rows.Add(300)
                .RowHeadersVisible = False
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                .Columns(0).HeaderText = "Dia [mu]"
                .Columns(1).HeaderText = "PSD Cum wght[%]"
            End With


            For h = 1 To norm_log_dist_pdf.GetLength(0) - 1 'Fill line chart
                xx = h / 6        'Calulation step
                '------------- Datagridview Bell curve ----------------
                norm_log_dist_pdf(h, 0) = xx
                norm_log_dist_pdf(h, 1) = Numerics.Distributions.Weibull.PDF(shape, scale, xx) * 100.0  '[%] Weibull

                '------------- Datagridview S curve (Cumulatief) ---- 
                norm_log_dist_cdf(h, 0) = xx
                If CheckBox22.Checked Then
                    norm_log_dist_cdf(h, 1) = (1.0 - Numerics.Distributions.Weibull.CDF(shape, scale, xx)) * 100.0  '[%] Weibull
                Else
                    norm_log_dist_cdf(h, 1) = (Numerics.Distributions.Weibull.CDF(shape, scale, xx)) * 100.0  '[%] Weibull
                End If

                '------------- Copy to the data grids ----
                DataGridView5.Rows(h - 1).Cells(0).Value = Round(norm_log_dist_pdf(h, 0), 2)    'Bell Curve
                DataGridView5.Rows(h - 1).Cells(1).Value = Round(norm_log_dist_pdf(h, 1), 2)    'Bell Curve
                DataGridView7.Rows(h - 1).Cells(0).Value = Round(norm_log_dist_cdf(h, 0), 2)    'Cum weight
                DataGridView7.Rows(h - 1).Cells(1).Value = Round(norm_log_dist_cdf(h, 1), 2)    'Cum weight
                If Round(norm_log_dist_cdf(h, 1)) = 50 Then
                    DataGridView7.Item(1, h - 1).Style.BackColor = Color.Red
                End If
            Next h
        End If
        Draw_chart5_weibull()

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        Dim words As String()
        Dim separators As String() = {";"}

        With ComboBox6
            words = typ_distri(.SelectedIndex).Split(separators, StringSplitOptions.None)

            NumericUpDown17.Value = CDec(Trim(words(1))) 'shape
            NumericUpDown24.Value = CDec(Trim(words(2))) 'scale

            .BackColor = CType(IIf(.SelectedIndex > 0, Color.LightGreen, Color.White), Color)

        End With
        Calc_plot_Distribution()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        With DataGridView6
            For row = 0 To .Rows.Count - 1
                If CDbl(DataGridView7.Rows(row).Cells(1).Value) > 0 Then
                    .Rows(row).Cells(0).Value = DataGridView7.Rows(row).Cells(0).Value
                    .Rows(row).Cells(1).Value = DataGridView7.Rows(row).Cells(1).Value
                Else
                    .Rows(row).Cells(0).Value = 0
                    .Rows(row).Cells(1).Value = 0
                End If
            Next
        End With
    End Sub
End Class
