VERSION 5.00
Begin VB.Form dialog_ExportHEIF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13110
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "File_Save_HEIF.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   874
   Begin PhotoDemon.pdButtonStrip btsCategory 
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1085
      FontSize        =   11
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5790
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4695
      Index           =   0
      Left            =   5880
      Top             =   1080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdButtonStrip btsQuality 
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1931
         Caption         =   "quality"
      End
      Begin PhotoDemon.pdSlider sldQuality 
         Height          =   405
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   6975
         _ExtentX        =   7223
         _ExtentY        =   873
         Max             =   100
         Value           =   90
         NotchPosition   =   2
         NotchValueCustom=   90
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   3
         Left            =   2880
         Top             =   2280
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "high quality, large file"
         FontItalic      =   -1  'True
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdLabel lblHint 
         Height          =   255
         Index           =   2
         Left            =   480
         Top             =   2280
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   450
         Caption         =   "low quality, small file"
         FontItalic      =   -1  'True
         FontSize        =   9
         ForeColor       =   4210752
         Layout          =   1
      End
      Begin PhotoDemon.pdCheckBox chkLivePreview 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         Caption         =   "preview quality changes"
         FontSize        =   11
      End
      Begin PhotoDemon.pdButtonStrip btsMultipage 
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1931
         Caption         =   "page format"
      End
   End
   Begin PhotoDemon.pdContainer picContainer 
      Height          =   4695
      Index           =   1
      Left            =   5880
      Top             =   1080
      Width           =   7215
      _ExtentX        =   0
      _ExtentY        =   0
      Begin PhotoDemon.pdMetadataExport mtdManager 
         Height          =   4215
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7435
      End
   End
End
Attribute VB_Name = "dialog_ExportHEIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'HEIF Export Dialog
'Copyright 2024-2025 by Tanner Helland
'Created: 26/August/24
'Last updated: 29/August/24
'Last update: add support for multi-page export (one layer per frame in the heif file)
'
'Dialog for presenting the user various options related to HEIF exporting.  All export options rely on
' libheif for their actual implementation.
'
'Unless otherwise noted, all source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file, or at https://photodemon.org/license/
'
'***************************************************************************

Option Explicit

'This form can (and should!) be notified of the image being exported.  The only exception to this rule is invoking
' the dialog from the batch process dialog, as no image is associated with that preview.
Private m_SrcImage As pdImage

'A composite of the current image, 32-bpp, fully composited.  This is only regenerated if the source image changes.
Private m_CompositedImage As pdDIB

'Clean, modified source preview image, in PNG format.  (This only needs to be created when the preview source
' image changes - e.g. if the user zooms or scrolls the preview control.)
Private m_PreviewImagePath As String, m_PreviewImageBackup As pdDIB

'OK or CANCEL result
Private m_UserDialogAnswer As VbMsgBoxResult

'Final format-specific XML packet, with all format-specific settings defined as tag+value pairs
Private m_FormatParamString As String

'Final metadata XML packet, with all metadata settings defined as tag+value pairs
Private m_MetadataParamString As String

'The user's answer is returned via this property
Public Function GetDialogResult() As VbMsgBoxResult
    GetDialogResult = m_UserDialogAnswer
End Function

Public Function GetFormatParams() As String
    GetFormatParams = m_FormatParamString
End Function

Public Function GetMetadataParams() As String
    GetMetadataParams = m_MetadataParamString
End Function

Private Sub btsCategory_Click(ByVal buttonIndex As Long)
    UpdatePanelVisibility
End Sub

Private Sub chkLivePreview_Click()
    UpdatePreview
End Sub

Private Sub UpdatePanelVisibility()
    Dim i As Long
    For i = 0 To btsCategory.ListCount - 1
        picContainer(i).Visible = (i = btsCategory.ListIndex)
    Next i
End Sub

Private Sub btsQuality_Click(ByVal buttonIndex As Long)
    UpdateQualityVisibility
    UpdatePreview
End Sub

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Hide
End Sub

Private Sub cmdBar_OKClick()
    
    'Store all parameters inside an XML string
    m_FormatParamString = GetParamString_HEIF()
    
    'The metadata panel manages its own XML string
    m_MetadataParamString = mtdManager.GetMetadataSettings
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    If (LenB(m_PreviewImagePath) > 0) Then Files.FileDeleteIfExists m_PreviewImagePath
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False
    
End Sub

Private Function GetParamString_HEIF() As String
    
    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "heif-lossless", (btsQuality.ListIndex = 0)
    cParams.AddParam "heif-lossy-quality", sldQuality.Value
    cParams.AddParam "heif-multiframe", (btsMultipage.ListIndex = 1), True
    
    GetParamString_HEIF = cParams.GetParamString()
    
End Function

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    
    sldQuality.Value = 90       'Arbitrary default
    btsQuality.ListIndex = 0    'Default to lossless mode
    
    'Default metadata settings
    mtdManager.Reset
    
End Sub

Private Sub Form_Load()
    chkLivePreview.AssignTooltip "This image format is very computationally intensive.  On older or slower PCs, you may want to disable live previews."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    'Ensure we release any temp files on exit
    Files.FileDeleteIfExists m_PreviewImagePath
    
    'Release subclassing form themer
    ReleaseFormTheming Me
    
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Make a copy of the composited image; it takes time to composite layers, so we don't want to redo this except
    ' when absolutely necessary.
    Set m_SrcImage = srcImage
    If ((m_SrcImage Is Nothing) Or (Not Plugin_Heif.IsLibheifEnabled())) Then
        Interface.ShowDisabledPreviewImage pdFxPreview
    Else
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    
    'Populate the category button strip
    btsCategory.AddItem "image", 0
    btsCategory.AddItem "metadata", 1
    
    'Populate the "image" options panel
    btsQuality.AddItem "lossless", 0
    btsQuality.AddItem "lossy", 1
    btsQuality.ListIndex = 0
    UpdateQualityVisibility
    
    'Single- or multi-frame export options (text reused from the TIFF export dialog)
    btsMultipage.AddItem "single page (composited image)", 0
    btsMultipage.AddItem "multipage (one page per layer)", 1
    btsMultipage.ListIndex = 0
    
    If (Not srcImage Is Nothing) Then btsMultipage.Visible = (srcImage.GetNumOfLayers > 1) Else btsMultipage.Visible = True
    
    'Next, prepare various controls on the metadata panel
    Set m_SrcImage = srcImage
    mtdManager.SetParentImage m_SrcImage, PDIF_HEIF
    
    'By default, the image options panel is always shown.
    btsCategory.ListIndex = 0
    UpdatePanelVisibility
    
    'Update the preview
    UpdatePreviewSource
    UpdatePreview True
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "HEIC")
    If (Not g_WindowManager Is Nothing) Then g_WindowManager.SetFocusAPI cmdBar.hWnd
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True

End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreviewSource
    UpdatePreview
End Sub

Private Sub sldQuality_Change()
    UpdatePreview
End Sub

'The dialog differentiates between lossless and lossy using a hard toggle; lossy settings are hidden
' when lossless mode is requested.
Private Sub UpdateQualityVisibility()
    
    Dim showLossySettings As Boolean
    showLossySettings = (btsQuality.ListIndex > 0)
    
    sldQuality.Visible = showLossySettings
    lblHint(2).Visible = showLossySettings
    lblHint(3).Visible = showLossySettings
    
End Sub

'When a parameter changes that requires a new base image for the preview (e.g. changing the background composite color),
' call this function to generate a new preview DIB.  Note that you *do not* need to call this function for format-specific
' changes (like quality, subsampling, etc).
Private Sub UpdatePreviewSource()

    If Not (m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region hasn't changed
        ' since the last preview.  Prep a new preview now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, False
        
        If (m_PreviewImageBackup Is Nothing) Then Set m_PreviewImageBackup = New pdDIB
        m_PreviewImageBackup.CreateFromExistingDIB workingDIB
        
    End If
    
End Sub

Private Sub UpdatePreview(Optional ByVal forceUpdate As Boolean = False)

    Const funcName As String = "UpdatePreview"
    
    If (Not Plugin_Heif.IsLibheifEnabled()) Then
        InternalError funcName, "libheif broken"
        Exit Sub
    End If
    
    If ((cmdBar.PreviewsAllowed Or forceUpdate) And (Not m_SrcImage Is Nothing)) Then
        
        'Make sure the preview source is up-to-date
        If (workingDIB Is Nothing) Then UpdatePreviewSource
        If (workingDIB Is Nothing) Then Exit Sub
        
        'Because previews are intensive to generate, this dialog provides a toggle so the user
        ' can suspend real-time previews.
        If chkLivePreview.Value Then
            Plugin_Heif.PreviewHEIF m_PreviewImageBackup, workingDIB, GetParamString_HEIF()
            FinalizeNonstandardPreview pdFxPreview, True
            
        'Live previews are disabled; just mirror the original image to the screen
        Else
            workingDIB.CreateFromExistingDIB m_PreviewImageBackup
            If (Not workingDIB.GetAlphaPremultiplication) Then workingDIB.SetAlphaPremultiplication True
            FinalizeNonstandardPreview pdFxPreview, True
        End If
    
    '/no else required, previews are disabled due to a valid reason (settings haven't changed, batch process, etc)
    End If

End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errMsg As String)
    PDDebug.LogAction "WARNING! Problem in dialog_ExportHEIF." & funcName & ": " & errMsg
End Sub
