VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DataEnvironment1 
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8685
   _ExtentX        =   15319
   _ExtentY        =   10557
   FolderFlags     =   1
   TypeLibGuid     =   "{85BEB2DA-1760-4D1B-B24B-CE22D74AA732}"
   TypeInfoGuid    =   "{4FA96BB2-7B60-429F-8737-10C75915FDB8}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Connection1"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Key\data_skor.mdb;Persist Security Info=False"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   1
   BeginProperty Recordset1 
      CommandName     =   "tabel_skor"
      CommDispId      =   1002
      RsDispId        =   1004
      CommandText     =   "tabel_skor"
      ActiveConnectionName=   "Connection1"
      CommandType     =   2
      dbObjectType    =   1
      Locktype        =   3
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "skor"
         Caption         =   "skor"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "nama"
         Caption         =   "nama"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DataEnvironment1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataEnvironment_Initialize()
    DataEnvironment1.Connection1.Open App.Path & "\data_skor.mdb"
End Sub
