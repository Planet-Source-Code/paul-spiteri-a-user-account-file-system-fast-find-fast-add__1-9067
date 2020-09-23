Attribute VB_Name = "RAF"
Option Explicit

'Paul Spiteri 's (paul@spiteri.com) User Account File
'Feel free to email me.
'Version 1.2
'
'Instructions:
'******************************************
'*REMEMBER TO CALL RAFCLEAR BEFORE USE!!!!*
'******************************************
'Put 'RAFStartup' in Form_Load
'MaxRecords will need to be changed to the maximum number of records. Default = 10000
''RAFClear' will need to be called once, to set up the pointer file.
'Call 'RAFAdd' to add a record.
'Call 'RAFSearch' to return a record. Pass the ID.
'Call 'RAFDelete" and pass the ID to delete, to delete its record.

' Change this type's fields for your needs.
Type User
    ID As Integer
    EMail As String * 30
    Password As String * 15
End Type

' Do not change this.
Type RAFRecord
    RecordNumber As Integer
End Type

Global UActive As Integer
Global UFile As String
Global ULength As Integer

Global RFile As String
Global RLength As Integer

' Set this in RAFStartup to the maximum number. 10000 is good.
Global MaxRecords As Integer

Public Sub RAFClear()
    Dim RChannel As Integer
    Dim i As Integer
    Dim TempRAF As RAFRecord
    
    TempRAF.RecordNumber = -1
    RChannel = FreeFile
    Open RFile For Random As RChannel Len = RLength
    For i = 1 To MaxRecords
        Put RChannel, i, TempRAF
    Next i
    Close RChannel
    Kill UFile
    
End Sub

Public Sub RAFStartup()
    Dim UChannel As Integer
    Dim RChannel As Integer
    Dim TempRAF As RAFRecord
    Dim User As User
    
    UFile = App.Path & "\Users.dat"
    RFile = App.Path & "\UPointer.dat"
    
    
    'UFile = "C:\Users.dat"
    'RFile = "C:\UPointer.dat"
    
    ULength = Len(User)
    RLength = Len(TempRAF)
    
    UChannel = FreeFile
    Open UFile For Random As UChannel Len = ULength
    Close UChannel
    
    RChannel = FreeFile
    Open RFile For Random As RChannel Len = RLength
    Close RChannel
    
    UActive = (FileLen(UFile) / ULength) + 1
    
    MaxRecords = 10000
    
    Form1.Text1.Text = UActive
End Sub

' The parameters will need to be changed when you change your User type.
Public Function RAFAdd(ID As Integer, AddEMail As String, AddPassword As String) As Boolean
    Dim UChannel As Integer
    Dim RChannel As Integer
    Dim TempRAF As RAFRecord
    Dim NewRAF As RAFRecord
    Dim User As User
    Dim RecordPosition As Integer
    Dim HashPos As Integer
    Dim Done As Boolean
    Dim Success As Boolean
    Dim StartPos As Integer

    'Add your parameters to the User here.
    User.ID = ID
    User.EMail = AddEMail
    User.Password = AddPassword
    
    RecordPosition = UActive
    NewRAF.RecordNumber = RecordPosition


    If RAFSearch(ID).ID = ID Then 'Checks if ID is already in use
        RAFAdd = False
    Else
        ' finds a free record position to put in pointer file
        HashPos = (ID Mod MaxRecords) + 1
        StartPos = HashPos
        Success = True
        Done = False
        
        RChannel = FreeFile
        Open RFile For Random As RChannel Len = RLength
        Do While Done = False
            Get RChannel, HashPos, TempRAF
            If TempRAF.RecordNumber = -1 Then
                Done = True
            Else
                If HashPos = MaxRecords Then
                    HashPos = 1
                Else
                    HashPos = HashPos + 1
                End If
                If HashPos = StartPos Then 'gone through file completely, no free slots.
                    Done = True
                    Success = False
                End If
            End If
        Loop
        Close RChannel
        
        If Success Then
            ' adds to the pointer file
            RChannel = FreeFile
            Open RFile For Random As RChannel Len = RLength
            Put RChannel, HashPos, NewRAF
            Close RChannel
            
            ' adds to the data file
            UChannel = FreeFile
            Open UFile For Random As UChannel Len = ULength
            Put UChannel, RecordPosition, User
            Close UChannel
            
            UActive = UActive + 1
            RAFAdd = True
        Else
            ' pointer fill must be full
            RAFAdd = False
        End If
    End If
    
End Function

Public Function RAFSearch(ID As Integer, Optional ByRef RecordPosition As Integer) As User
    Dim UChannel As Integer
    Dim RChannel As Integer
    Dim TempRAF As RAFRecord
    Dim User As User
    Dim HashPos As Integer
    Dim Done As Boolean
    Dim Success As Boolean
    Dim StartPos As Integer

    HashPos = (ID Mod MaxRecords) + 1
    StartPos = HashPos
    Success = True
    Done = False
    
    RChannel = FreeFile
    Open RFile For Random As RChannel Len = RLength
    UChannel = FreeFile
    Open UFile For Random As UChannel Len = ULength
    Do While Done = False
        Get RChannel, HashPos, TempRAF
        If TempRAF.RecordNumber <> -1 And TempRAF.RecordNumber <> 0 Then
            Get UChannel, TempRAF.RecordNumber, User
        End If
        If User.ID = ID Then
            Done = True
        Else
            If HashPos = MaxRecords Then
                HashPos = 1
            Else
                HashPos = HashPos + 1
            End If
            If HashPos = StartPos Then
                Done = True
                Success = False
            End If
        End If
    Loop
    Close RChannel
    Close UChannel
    
    If Not Success Then
        User.ID = -1
        RecordPosition = -1
    Else
        RecordPosition = HashPos
    End If
    RAFSearch = User
End Function

Public Function RAFDelete(ID As Integer) As Boolean
    Dim RChannel As Integer
    Dim TempRAF As RAFRecord
    Dim RecordPos As Integer

    RAFSearch ID, RecordPos

    If RecordPos <> -1 Then
        TempRAF.RecordNumber = -1
        RChannel = FreeFile
        Open RFile For Random As RChannel Len = RLength
        Put RChannel, RecordPos, TempRAF
        Close RChannel
        RAFDelete = True
    Else
        RAFDelete = False
    End If
End Function

