VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "SWFLASH.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form2 
   Caption         =   "Zyweb"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   11880
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7125
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   11895
      ExtentX         =   20981
      ExtentY         =   12568
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1272
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar3 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   7845
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   2910
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&On The Web"
            Key             =   "Onthewb"
            Description     =   "Misc Websites"
            Object.ToolTipText     =   "Misc. Web-Sites Of Intrest"
            Object.Tag             =   "Onthewb"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   9
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Moma"
                  Object.Tag             =   "Moma"
                  Text            =   "Museum of Modern Art"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Pets"
                  Object.Tag             =   "Pets"
                  Text            =   "PETS.COM"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Staples"
                  Object.Tag             =   "Staples"
                  Text            =   "Staples.com"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PlanetRX"
                  Object.Tag             =   "PlanetRX"
                  Text            =   "PlanetRx.com"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Garden"
                  Object.Tag             =   "Garden"
                  Text            =   "Garden.com"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BandN"
                  Object.Tag             =   "BandN"
                  Text            =   "Barnes And Noble"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Shock"
                  Object.Tag             =   "Shock"
                  Text            =   "Shockwave"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Mw"
                  Object.Tag             =   "Mw"
                  Text            =   "Micro Warehouse"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BePaid"
                  Object.Tag             =   "BePaid"
                  Text            =   "Be Paid To Surf The Web"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Save Document"
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save Web-Page To Disk"
            Object.Tag             =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&New Window"
            Key             =   "New"
            Description     =   "Open A New Window"
            Object.ToolTipText     =   "Open A New Window"
            Object.Tag             =   "New"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Web Radio"
            Key             =   "Webradio"
            Description     =   "Tune Into The Web-Radio"
            Object.ToolTipText     =   "Tune Into The Web-Radio"
            Object.Tag             =   "webradio"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   9500
         TabIndex        =   5
         Top             =   40
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":377A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4602
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5886
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":6B0A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   2011
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Favorites"
            Key             =   "Favorites"
            Description     =   "Favorites Menu"
            Object.ToolTipText     =   "Favorites Menu"
            Object.Tag             =   "Favories"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Add"
                  Object.Tag             =   "Add"
                  Text            =   "&Add Favorite"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Edit"
                  Object.Tag             =   "Edit"
                  Text            =   "&Edit Favorites"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Spacer"
                  Object.Tag             =   "Spacer"
                  Text            =   "-"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   11570
         Picture         =   "Form2.frx":7D8E
         ScaleHeight     =   210
         ScaleWidth      =   240
         TabIndex        =   10
         Top             =   50
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   23
         Width           =   8895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "G&o"
         Default         =   -1  'True
         Height          =   255
         Left            =   10440
         TabIndex        =   2
         Top             =   50
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   2011
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Navigate"
            Key             =   "Navigate"
            Description     =   "Navigate To A Web-Site"
            Object.ToolTipText     =   "Navigate To A Selected Web-Site"
            Object.Tag             =   "Navigate"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Stop"
            Key             =   "Stop"
            Description     =   "Halt Current Transfer"
            Object.ToolTipText     =   "Halt Current Transfer."
            Object.Tag             =   "Stop"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Reload"
            Key             =   "Reload"
            Description     =   "Reload Current Web-Page"
            Object.ToolTipText     =   "Reload Current Web-Page."
            Object.Tag             =   "Reload"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Previous"
            Key             =   "Previous"
            Description     =   "Previous Web-Page"
            Object.ToolTipText     =   "Previous Web-Page."
            Object.Tag             =   "Previous"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Next"
            Key             =   "Next"
            Description     =   "Next Webpage"
            Object.ToolTipText     =   "Next Web-Page"
            Object.Tag             =   "Next"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
            Key             =   "Find"
            Description     =   "Find Web-Page"
            Object.ToolTipText     =   "Find A Web-Page."
            Object.Tag             =   "Find"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Home"
            Key             =   "Home"
            Description     =   "Home-Page"
            Object.ToolTipText     =   "Navigate To Homepage"
            Object.Tag             =   "Home"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Print"
            Key             =   "Print"
            Description     =   "Print Current Web-Page"
            Object.ToolTipText     =   "Print Current Web-Page"
            Object.Tag             =   "Print"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Options"
            Key             =   "Options"
            Description     =   "Options"
            Object.ToolTipText     =   "Edit Options For This Program"
            Object.Tag             =   "Options"
            ImageIndex      =   9
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
         Height          =   375
         Left            =   11520
         TabIndex        =   4
         Top             =   0
         Width           =   370
         _cx             =   4194957
         _cy             =   4194965
         Movie           =   "c:\my documents\loading.swf"
         Src             =   "c:\my documents\loading.swf"
         WMode           =   "Window"
         Play            =   0   'False
         Loop            =   -1  'True
         Quality         =   "High"
         SAlign          =   "L"
         Menu            =   -1  'True
         Base            =   ""
         Scale           =   "ShowAll"
         DeviceFont      =   0   'False
         EmbedMovie      =   -1  'True
         BGColor         =   ""
         SWRemote        =   ""
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":7E21
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":90A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":9F2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":A381
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":B209
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":C091
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":D315
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":E19D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":F025
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar4 
      Align           =   4  'Align Right
      Height          =   7125
      Left            =   11520
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   12568
      ButtonWidth     =   2514
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList4"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close Window"
            Key             =   "Close"
            Description     =   "Close Window"
            Object.ToolTipText     =   "Close Popup Window"
            Object.Tag             =   "Close"
            ImageIndex      =   1
         EndProperty
      EndProperty
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   1335
         Left            =   45
         TabIndex        =   9
         Top             =   360
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2355
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   0
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":FEAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":10D35
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblReturn As Double
Private Sub Command1_Click()
'Navigate To Selected Web-page
Me.WebBrowser1.Navigate Text1.Text
End Sub
Private Sub Form_Load()
'WebBrowser1.Navigate2 "http://www.3rdforce.com/swa/moon.html"
End Sub
Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then Exit Sub
If WebBrowser1.MenuBar = False Then Exit Sub
Me.Height = 8580
Me.Width = 12000
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Subclasing For The Toolbar
On Error Resume Next
Select Case Button.Key
Case "Navigate"
If Text1.Text = "" Then
MsgBox "You Must Enter A URL"
Exit Sub
End If
WebBrowser1.Navigate2 Text1.Text
Case "stop"
WebBrowser1.Stop
Case "reload"
WebBrowser1.Refresh
Case "Previous"
WebBrowser1.GoBack
Case "Next"
WebBrowser1.GoForward
Case "Find"
WebBrowser1.GoSearch
Case "home"
WebBrowser1.Navigate "http://homepage.hitter.net/spooner/modis"
Case "Print"
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
Case "Options"
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Select
End Sub
Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case "1"
MsgBox "Welcome To On The Web Click The Arrow To The Right For A List Of Pages On The Web"
Case "2"
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER
Case "3"
Dim frmWB As Form2
Set frmWB = New Form2
frmWB.Visible = True
Case "4"
WebBrowser1.Navigate "http://radio.sonicnet.com/mymusiclisten.asp?name=Secure2k"
Case "5"
'WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_PROMPTUSER
End Select
End Sub
Private Sub Toolbar3_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error Resume Next
Select Case ButtonMenu.Index
 Case "1"
 WebBrowser1.Navigate "http://service.bfast.com/bfast/click?bfmid=6866694&siteid=13245143&bfpage=moma_storefront"
 Case "2"
 WebBrowser1.Navigate "http://service.bfast.com/bfast/click?bfmid=1946520&siteid=13245172&bfpage=home"
 Case "3"
 WebBrowser1.Navigate "http://service.bfast.com/bfast/click?bfmid=573112&siteid=13245259&bfpage=home012"
 Case "4"
 WebBrowser1.Navigate "http://service.bfast.com/bfast/click?bfmid=1000057&siteid=13318489&bfpage=homepage"
 Case "5"
 WebBrowser1.Navigate "http://service.bfast.com/bfast/click?bfmid=9664445&siteid=14296546&bfpage=home"
 Case "6"
 WebBrowser1.Navigate "http://service.bfast.com/bfast/click?bfmid=2181&sourceid=13245364&categoryid=rn_home"
 Case "7"
 WebBrowser1.Navigate "http://service.bfast.com/bfast/click?bfmid=7362327&siteid=13245188&bfpage=general"
 Case "8"
 WebBrowser1.Navigate "http://service.bfast.com/bfast/click?bfmid=1391718&siteid=13245217&bfpage=homepage"
 Case "9"
 WebBrowser1.Navigate "http://www.bepaid.com/master.rhtml?REFID=10595823"
 End Select
End Sub
Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case "1"
Me.Hide
End Select
End Sub
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
ShockwaveFlash1.Playing = False
ShockwaveFlash1.Rewind
End Sub
Private Sub WebBrowser1_DownloadBegin()
On Error Resume Next
Me.Caption = "Zyweb"
Me.Text1.Text = WebBrowser1.LocationURL
ShockwaveFlash1.Playing = True
End Sub
Private Sub WebBrowser1_DownloadComplete()
Me.Caption = "Zyweb - " & WebBrowser1.LocationName
On Error Resume Next
If WebBrowser1.MenuBar = False Then
Me.BorderStyle = 3
Me.Height = WebBrowser1.Height
Me.Width = WebBrowser1.Width
Me.Height = Me.Height + 375
Me.Width = Me.Width + 80 + 400
Toolbar1.Visible = False
Toolbar2.Visible = False
Toolbar3.Visible = False
Toolbar4.Visible = True
If Me.WindowState = vbMinimized Then Exit Sub
WebBrowser1.Left = 0
WebBrowser1.Top = 0
End If
ShockwaveFlash1.Rewind
ShockwaveFlash1.Playing = False

End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Dim frmWB As Form2
Set frmWB = New Form2
frmWB.WebBrowser1.RegisterAsBrowser = True
Set ppDisp = frmWB.WebBrowser1.Object
frmWB.Visible = True
End Sub
Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
If Progress = -1 Then ProgressBar1.Value = 100 & ProgressBar2.Value = 100
If Progress > 0 And ProgressMax > 0 Then
ProgressBar1.Value = Progress * 100 / ProgressMax
ProgressBar2.Value = ProgressBar1.Value
End If
End Sub
Private Sub WebBrowser1_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
On Error Resume Next
Me.Visible = False
End Sub
