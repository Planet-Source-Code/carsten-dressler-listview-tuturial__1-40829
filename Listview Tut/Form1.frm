VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "ListView Tut 1.0"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Add Stuff"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   4440
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   0
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
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8493
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SubItem 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SubItem 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SubItem 3"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Carsten Dressler - IntraDream.com"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label Label7 
      Caption         =   "6. Now the the Coding.....Click the button to see how it works....btw its commented :)"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label Label6 
      Caption         =   "5.Right Click the Listview and in the ImageList Tab select the SMALL ImageList ComboBox and select the imagelist that is aviable."
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   3735
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":039A
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "3.Head on Over to the Colume Header Section...and add as many Headers as u like."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "2.Right Click it, then in the General  Tab Select the View ComboBox then select  3-LvwReport"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "1.Add ListView to a form!"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "ListView Tuturial for PSCODE Visitors"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   'Adds a Item to the ListView
   'This is how the Listitems.Add is ordered - Index, Key, Text, Icon, SmallIcon
   'The 11 is the icon that is selected for the smallIcon...To find out what icon is what color right click the imagelist and select the icon it will tell u
   ListView1.ListItems.Add , , "Item 1", , 1
   
   'This changes the Subitems of the created item
   ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = "Sub1"
   ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = "Sub2"
   
  'This is how you would loop though a Listview
  'Dim lngA as long
  'For lngA = 1 to listview1.listitems.count
  '   Dim A as string
  '   A = listview1.listitems(lngA).text  'or Subitems(2)
  'Next lnga
  
  'Isnt this shit simple?
  
   
   
End Sub
