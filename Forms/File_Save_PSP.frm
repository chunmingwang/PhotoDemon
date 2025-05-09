VERSION 5.00
Begin VB.Form dialog_ExportPSP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12630
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
   Icon            =   "File_Save_PSP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   439
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   Begin PhotoDemon.pdSlider sldCompression 
      Height          =   855
      Left            =   5880
      TabIndex        =   3
      Top             =   1440
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1508
      Caption         =   "compression level"
      Max             =   12
      Value           =   9
      NotchPosition   =   2
      NotchValueCustom=   9
   End
   Begin PhotoDemon.pdCommandBar cmdBar 
      Align           =   2  'Align Bottom
      Height          =   750
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   1323
      DontAutoUnloadParent=   -1  'True
   End
   Begin PhotoDemon.pdFxPreviewCtl pdFxPreview 
      Height          =   5625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9922
   End
   Begin PhotoDemon.pdButtonStrip btsCompatibility 
      Height          =   1095
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1931
      Caption         =   "target version"
   End
End
Attribute VB_Name = "dialog_ExportPSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Corel/JASC Paintshop Pro (PSP) Export Dialog
'Copyright 2021-2025 by Tanner Helland
'Created: 05/February/21
'Last updated: 05/February/21
'Last update: wrap up initial build
'
'This dialog works as a simple relay to the pdPSP class (and its associated child classes).
' Look there for specific encoding details.
'
'I have tried to pare down the UI toggles to only the most essential elements.  Most PSP-compatible
' settings will be automatically generated by PD, where applicable - the caller just needs to
' specify compression level (which greatly affects export time, at a trade-off to final file size)
' and a "target compatibility" setting (which is relevant because older PSP versions don't support
' some features that PD does).
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

Private Sub cmdBar_CancelClick()
    m_UserDialogAnswer = vbCancel
    Me.Visible = False
End Sub

Private Sub cmdBar_OKClick()

    Dim cParams As pdSerialize
    Set cParams = New pdSerialize
    cParams.AddParam "compression-level", sldCompression.Value, True
    
    Dim strCompatTarget As String
    If (btsCompatibility.ListIndex = 0) Then
        strCompatTarget = "auto"
    Else
        strCompatTarget = btsCompatibility.ListIndex + 5  'Convert .ListIndex to [6, 8] scale
    End If
    
    cParams.AddParam "compatibility-target", strCompatTarget
    
    m_FormatParamString = cParams.GetParamString
    
    'Metadata export is not currently supported
    m_MetadataParamString = vbNullString
    
    'Free resources that are no longer required
    Set m_CompositedImage = Nothing
    Set m_SrcImage = Nothing
    
    'Hide but *DO NOT UNLOAD* the form.  The dialog manager needs to retrieve the setting strings before unloading us
    m_UserDialogAnswer = vbOK
    Me.Visible = False

End Sub

Private Sub cmdBar_RequestPreviewUpdate()
    UpdatePreview
End Sub

Private Sub cmdBar_ResetClick()
    sldCompression.Value = 9
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReleaseFormTheming Me
End Sub

Private Sub pdFxPreview_ViewportChanged()
    UpdatePreview
End Sub

'The ShowDialog routine presents the user with this form.
Public Sub ShowDialog(Optional ByRef srcImage As pdImage = Nothing)
    
    'Provide a default answer of "cancel" (in the event that the user clicks the "x" button in the top-right)
    m_UserDialogAnswer = vbCancel
    
    'Make sure that the proper cursor is set
    Screen.MousePointer = 0
    Message "Waiting for user to specify export options... "
    
    'Populate any list elements
    btsCompatibility.AddItem "auto", 0
    btsCompatibility.AddItem "6", 1
    btsCompatibility.AddItem "7", 2
    btsCompatibility.AddItem "8+", 3
    btsCompatibility.ListIndex = 0
    
    'Prep a preview (if any)
    Set m_SrcImage = srcImage
    If (Not m_SrcImage Is Nothing) Then
        m_SrcImage.GetCompositedImage m_CompositedImage, True
        pdFxPreview.NotifyNonStandardSource m_CompositedImage.GetDIBWidth, m_CompositedImage.GetDIBHeight
    End If
    If (m_SrcImage Is Nothing) Then Interface.ShowDisabledPreviewImage pdFxPreview
    
    UpdatePreview
    
    'Apply translations and visual themes
    ApplyThemeAndTranslations Me, True, True
    Interface.SetFormCaptionW Me, g_Language.TranslateMessage("%1 options", "PSP")
    
    'Display the dialog
    ShowPDDialog vbModal, Me, True
    
End Sub

Private Sub UpdatePreview()

    If cmdBar.PreviewsAllowed And (Not m_SrcImage Is Nothing) And (Not m_CompositedImage Is Nothing) Then
        
        'Because the user can change the preview viewport, we can't guarantee that the preview region
        ' hasn't changed since the last preview.  Prep a new preview base image now.
        Dim tmpSafeArray As SafeArray2D
        EffectPrep.PreviewNonStandardImage tmpSafeArray, m_CompositedImage, pdFxPreview, True
        EffectPrep.FinalizeNonstandardPreview pdFxPreview, True
        
    End If

End Sub
