Imports System.Runtime.InteropServices
Imports System.Windows.Interop

Friend Enum AccentState
	ACCENT_DISABLED = 0
	ACCENT_ENABLE_GRADIENT = 1
	ACCENT_ENABLE_TRANSPARENTGRADIENT = 2
	ACCENT_ENABLE_BLURBEHIND = 3
	ACCENT_INVALID_STATE = 4
End Enum

Friend Structure AccentPolicy
    Dim AccentState As AccentState
    Dim AccentFlags As Integer
    Dim GradientColor As Integer
    Dim AnimationId As Integer
End Structure

Friend Structure WindowCompositionAttributeData
    Dim Attribute As WindowCompositionAttribute
    Dim Data As IntPtr
    Dim SizeOfData As Integer
End Structure

Friend Enum WindowCompositionAttribute
	' ...
	WCA_ACCENT_POLICY = 19
	' ...
End Enum

''' <summary>
''' Interaction logic for MainWindow.xaml
''' </summary>
Partial Public Class MainWindow
    Inherits Window

    Friend Declare Function SetWindowCompositionAttribute Lib "user32.dll" (hwnd As IntPtr, ByRef data As WindowCompositionAttributeData) As Integer

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        EnableBlur()
    End Sub

    Friend Sub EnableBlur()
        Dim windowHelper As New WindowInteropHelper(Me)

        Dim accent As New AccentPolicy With {
            .AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND
        }

        Dim hGc = GCHandle.Alloc(accent, GCHandleType.Pinned)
        Dim data As New WindowCompositionAttributeData With {
            .Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY,
            .SizeOfData = Marshal.SizeOf(accent),
            .Data = hGc.AddrOfPinnedObject
        }
        hGc.Free()

        SetWindowCompositionAttribute(windowHelper.Handle, data)
    End Sub

End Class