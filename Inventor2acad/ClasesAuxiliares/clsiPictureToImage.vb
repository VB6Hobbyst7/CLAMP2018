Imports stdole

Public Class clsiPictureToImage
    Inherits System.Windows.Forms.AxHost

    Public Sub New()
        MyBase.New("{63109182-966B-4e3c-A8B2-8BC4A88D221C}")
        'New(Nothing)
    End Sub

    Public Function GetImageFromIPictureDisp(ByVal pictureDisp As System.Object) As System.Drawing.Image
        Dim objPicture As System.Drawing.Image
        'objPicture = CType(Windows.Forms.AxHost.GetPictureFromIPicture(pictureDisp), System.Drawing.Image)
        objPicture = CType(GetPictureFromIPicture(pictureDisp), System.Drawing.Image)

        Return objPicture
        'Return CType(Windows.Forms.AxHost.GetPictureFromIPicture(objImage), System.Drawing.Image)
    End Function

    Public Function GetIPictureDispFromImage(ByVal objImage As System.Drawing.Image) As System.Object
        'Dim objPicture As stdole.IPictureDisp
        'objPicture = CType(Windows.Forms.AxHost.GetIPictureDispFromPicture(objImage), stdole.IPictureDisp)
        Return DirectCast(GetIPictureDispFromPicture(objImage), System.Object)

        'Return objPicture
        'Return CType(Windows.Forms.AxHost.GetIPictureDispFromPicture(objImage), stdole.IPictureDisp)
    End Function
End Class


' Utility class that provides support for converting between
' IPictureDisp and Image objects.
Friend Class AxHostConverter
    Inherits System.Windows.Forms.AxHost

    Private Sub New()
        MyBase.New("")
    End Sub

    Public Shared Function ImageToPictureDisp( _
                    ByVal objImage As System.Drawing.Image) As System.Object
        Return DirectCast(GetIPictureDispFromPicture(objImage),  _
                          System.Object)
    End Function

    Public Shared Function PictureDispToImage( _
                  ByVal pictureDisp As stdole.IPictureDisp) As System.Drawing.Image
        Dim objPicture As System.Drawing.Image
        objPicture = CType(GetPictureFromIPicture(pictureDisp),  _
                           System.Drawing.Image)

        Return objPicture
    End Function
End Class
