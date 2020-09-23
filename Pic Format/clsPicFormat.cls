VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPicFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Max Pic Size: 34755 x 34755

Public Enum Quality
    Q1_Best = 1
    Q2_Excellent = 2
    Q3_Recommended = 3
    Q4_Good = 4
    Q5_Poor = 5
    Q6_Bad = 6
    Q7_Terrible = 7
End Enum
'Info About The Different Qualities
'Best        | 12.5% Smaller |  2097152 Colors | 21 Bits Per Pixel
'Excellent   | 25.0% Smaller |   262144 Colors | 18 Bits Per Pixel
'Recommended | 37.5% Smaller |    32768 Colors | 15 Bits Per Pixel
'Good        | 50.0% Smaller |     4096 Colors | 12 Bits Per Pixel
'Poor        | 62.5% Smaller |      512 Colors |  9 Bits Per Pixel
'Bad         | 75.0% Smaller |       64 Colors |  6 Bits Per Pixel
'Terrible    | 87.5% Smaller |        8 Colors |  3 Bits Per Pixel

'You Could Have A New Quality
'But It Would Just Be A Bmp File
'Perfect     | 00.0% Smaller | 16777216 Colors | 24 Bits Per Pixel

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Const FILE_SHARE_READ        As Long = &H1
Private Const FILE_SHARE_WRITE       As Long = &H2
Private Const GENERIC_WRITE          As Long = &H40000000
Private Const GENERIC_READ           As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL  As Long = &H80
Private Const CREATE_ALWAYS          As Long = 2
Private Const OPEN_ALWAYS            As Long = 4

Private PicInfo As BITMAPINFOHEADER
Private hBMO As Long, hMyDC As Long, hDIB As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Sub Create(W As Long, H As Long)
    'Creates An Off Screen Picture
    If H > 34755 Or W > 34755 Then
        MsgBox "Error Pic To Big.", vbExclamation, "Error!"
        Class_Terminate
        Exit Sub
    End If
    Class_Terminate
    PicInfo.biWidth = W
    PicInfo.biHeight = H
    PicInfo.biSizeImage = 3 * H * W
    hMyDC = CreateCompatibleDC(0)
    hDIB = CreateDIBSection(hMyDC, PicInfo, 0, 0, 0, 0)
    If hDIB Then
        hBMO = SelectObject(hMyDC, hDIB)
    Else
        MsgBox "Error: DIB failed to create" & vbNewLine & "Please restart your computer and try again.", vbExclamation, "Error!"
        End
    End If
End Sub
Private Sub FileRead(FileName As String, Data() As Byte)
    Dim hFile As Long, Size As Long, BytesRead As Long
    
    hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = -1 Then
        MsgBox "Error: Can't Open File", 0, "Error"
        Exit Sub
    End If
    Size = FileLen(FileName)
    ReDim Data(Size - 1) As Byte
    ReadFile hFile, Data(0), Size, BytesRead, 0
    FlushFileBuffers hFile
    CloseHandle hFile
End Sub
Private Sub FileWrite(FileName As String, Data() As Byte)
    Dim hFile As Long, Size As Long, BytesWritten As Long
    
    If Dir(FileName, 63) <> vbNullString Then Kill FileName
    hFile = CreateFile(FileName, GENERIC_WRITE, FILE_SHARE_WRITE, 0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = -1 Then
        MsgBox "Error: Can't Open File", 0, "Error"
        Exit Sub
    End If
    Size = UBound(Data) + 1
    WriteFile hFile, Data(0), Size, BytesWritten, 0
    FlushFileBuffers hFile
    CloseHandle hFile
End Sub
Public Sub LoadPic(FileName As String, hDestDC As Long)
    Dim A As Long, Data() As Byte, Temp As Long, PicData() As Byte
    Dim Num As Long, Num2 As Long, Pos As Long, Quality As Quality
    Dim PowersOfTwo(7) As Long, BitsPerColor As Long, BitsUsed As Long
    
    FileRead FileName, Data
    Quality = Data(4) Mod 8
    If Quality = 0 Then Exit Sub
    Create Data(0) * 256 + Data(1), Data(2) * 256 + Data(3)
    PowersOfTwo(0) = 1
    PowersOfTwo(1) = 2
    PowersOfTwo(2) = 4
    PowersOfTwo(3) = 8
    PowersOfTwo(4) = 16
    PowersOfTwo(5) = 32
    PowersOfTwo(6) = 64
    PowersOfTwo(7) = 128
    Num = PowersOfTwo(Quality)
    Num2 = PowersOfTwo(Quality - 1)
    'First 2 Bytes In Data Is The Width
    'Second 2 Bytes In Data Is The Height
    'The First 3 Bits In The 5th Byte Is The Quality
    'That's Why Pos = 4 And BitsUsed = 3
    Pos = 4
    BitsUsed = 3
    BitsPerColor = 8 - Quality
    ReDim PicData(PicInfo.biSizeImage - 1)
    For A = 0 To UBound(PicData)
        'Reads Data() Bit By Bit
        If Pos = UBound(Data) Then
            Temp = Data(Pos)
        Else
            Temp = CLng(Data(Pos + 1)) * 256 + Data(Pos)
        End If
        Temp = (Temp \ PowersOfTwo(BitsUsed)) Mod PowersOfTwo(BitsPerColor)
        'Sets One Of The Values (R,G,Or B) In The Picture To The Value Read From Data()
        PicData(A) = Temp * Num + Num2
        'Updates How Many Bits Have Been Read From Data(Pos)
        BitsUsed = BitsUsed + BitsPerColor
        If BitsUsed > 7 Then
            BitsUsed = BitsUsed - 8
            Pos = Pos + 1
        End If
    Next
    'You Can Only Go Directly From PicData() To A PictureBox
    'If The Screen Color Depth Was 24 bit When The PictureBox Was Created
    
    'Copies Data From PicData() To My Off Screen Picture
    SetBitmapBits hDIB, PicInfo.biSizeImage, PicData(0)
    'Copies Data From My Off Screen Picture To hDestDC
    BitBlt hDestDC, 0, 0, PicInfo.biWidth, PicInfo.biHeight, hMyDC, 0, 0, vbSrcCopy
End Sub
Public Sub SavePic(FileName As String, hSorDC As Long, W As Long, H As Long, Quality As Quality)
    Dim A As Long, Data() As Byte, PicData() As Byte
    Dim Num As Long, Pos As Long, Temp As Long
    Dim PowersOfTwo(7) As Long, BitsPerColor As Long, BitsUsed As Long
    
    If Quality < Q1_Best Or Quality > Q7_Terrible Then Exit Sub

    Create W, H
    ReDim PicData(PicInfo.biSizeImage - 1)
    'Copies Data From hDestDC To My Off Screen Picture
    BitBlt hMyDC, 0, 0, W, H, hSorDC, 0, 0, vbSrcCopy
    'Copies Data From My Off Screen Picture To PicData()
    GetBitmapBits hDIB, PicInfo.biSizeImage, PicData(0)
    'First 2 Bytes In Data Is The Width
    'Second 2 Bytes In Data Is The Height
    'The First 3 Bits In The 5th Byte Is The Quality
    'That's Why Pos = 4 And BitsUsed = 3
    Pos = 4
    BitsUsed = 3
    BitsPerColor = 8 - Quality
    'Calculate How Big The File Will Be
    Temp = 3 + BitsPerColor * PicInfo.biSizeImage
    Num = Temp \ 8
    If Temp Mod 8 Then Num = Num + 1
    ReDim Data(Num + 4)
    Data(0) = W \ 256
    Data(1) = W Mod 256
    Data(2) = H \ 256
    Data(3) = H Mod 256
    Data(4) = Quality

    PowersOfTwo(0) = 1
    PowersOfTwo(1) = 2
    PowersOfTwo(2) = 4
    PowersOfTwo(3) = 8
    PowersOfTwo(4) = 16
    PowersOfTwo(5) = 32
    PowersOfTwo(6) = 64
    PowersOfTwo(7) = 128
    Num = PowersOfTwo(Quality)
    For A = 0 To UBound(PicData)
        'Reads One Of The Values (R,G,Or B) In The Picture
        'And Reduces Color Writes It Bit By Bit Into Data()
        Temp = (PicData(A) \ Num) * PowersOfTwo(BitsUsed)
        Data(Pos) = Data(Pos) + (Temp Mod 256)
        'Updates How Many Bits Have Been Read From Data(Pos)
        BitsUsed = BitsUsed + BitsPerColor
        If BitsUsed > 7 Then
            BitsUsed = BitsUsed - 8
            Pos = Pos + 1
            Temp = Temp \ 256
            'If A File Is Using Every Bit Then This Will
            'Prevent It From Making An Out Of Range Error
            If Temp Then Data(Pos) = Data(Pos) + Temp
        End If
    Next
    FileWrite FileName, Data
End Sub
Private Sub Class_Initialize()
    PicInfo.biSize = Len(PicInfo)
    PicInfo.biPlanes = 1
    PicInfo.biBitCount = 24
End Sub
Private Sub Class_Terminate()
    'Deletes Off Screen Picture
    If hMyDC Then
        If hBMO Then
            DeleteObject SelectObject(hMyDC, hBMO)
            hBMO = 0
        End If
        DeleteObject hMyDC
        hMyDC = 0
    End If
    PicInfo.biWidth = 0
    PicInfo.biHeight = 0
    PicInfo.biSizeImage = 0
End Sub
