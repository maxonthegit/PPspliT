Attribute VB_Name = "PPspliT"
'
'
'    _____  _____           _ _ _______
'   |  __ \|  __ \         | (_)__   __|
'   | |__) | |__) |__ _ __ | |_   | |
'   |  ___/|  ___/ __| '_ \| | |  | |
'   | |    | |   \__ \ |_) | | |  | |
'   |_|    |_|   |___/ .__/|_|_|  |_|
'                    | |
'                    |_| by Massimo Rimondini - version 2.2
'
' first written by Massimo Rimondini in November 2009
' last update: December 2022
' Source code for PowerPoint 2007+
'
'




' This global variable indicates whether and how slide numbers should be kept
' consistent with the original set of slides. For example, if slide 6 is split
' into 3 slides, then all those 3 slides will be numbered 6 after splitting.
' As an alternative option, a subindex can be added to slide numbers, so that,
' for example, slide 6 is split into 6.1, 6.2, 6.3, etc.
Public slideNumbersAdjustMode As Integer
Public Const SLIDENUMBER_DONOTHING = 0
Public Const SLIDENUMBER_BAKE = 1
Public Const SLIDENUMBER_SUBINDEX = 2

' This global variable indicates whether animations should be split
' at each mouse-triggered event. If set to false, a separate slide is
' created for each and every animation.
Public splitMouseTriggered As Boolean

' The following variables are for internal use only.
Public cancelStatus As Boolean
Public slide_number As Integer

' Required at least for Office 2010 64 bit, as Windows Common Controls are
' not available
Public Const maxProgressWidth = 324

'
' Convert decimal separators in the argument string from '.' to the most
' appropriate character for the system-configured locale.
'
Private Function localizeDecimalSeparators(ByVal s As String)
    Dim d As Double, useCommaAsSeparator As Boolean
    useCommaAsSeparator = False
    
    ' Use a test value to check for the currently used decimal
    ' separator. In principle, we could use the user-supplied
    ' argument, but if it is a value between 0 and 1, it could
    ' miss the leading zero (e.g., -.1234), thus raising errors
    ' if we are not using the correct decimal separator in the
    ' assignment (which is exactly what we are trying to
    ' discover here).
    
    d = "1,2"
    ' If "," is not the decimal separator in use for the current
    ' system locale, this assignment results in losing the decimal
    ' separator.
    ' Now, this test requires care: in fact, localization of
    ' Double values seems to happen whenever a value is output on
    ' screen or is converted from a string, but in some way it does
    ' not seem to affect the internal representation of the Double
    ' value. Therefore, to check whether the decimal separator
    ' has survived the assignment, we need to look for its
    ' internal representation (which is "."), not its localized one.
    useCommaAsSeparator = (InStr(Trim(Str$(d)), ".") > 0)
    
    If useCommaAsSeparator Then
        d = Replace(s, ".", ",")
    Else
        d = s
    End If
    localizeDecimalSeparators = d
End Function

'
' This function looks for a shape in a slide by its ID (which has
' been previously copied to a tag).
' It returns Nothing if no such shape is found.
'
Private Function findShapeByIDinTag(shapeID As Long, slide As slide)
    Dim s As Shape
    
    Set findShapeByIDinTag = Nothing
    For Each s In slide.Shapes
        If s.Tags("ID") = Str$(shapeID) Then
            Set findShapeByIDinTag = s
            Exit Function
        End If
    Next s
End Function


'
' This utility sub is required to wrap the action of accessing
' property Brightness of ColorFormat objects, because this
' property is only exposed starting from Office 2010 (14.0).
' Since VBA's compiler evaluates the full code of a subroutine
' when it is about to execute it, putting this assignment in a
' conditional statement is not enough to prevent it from raising
' a "Method or data member not found" error at runtime.
'
Private Sub assignColorBrightness(col1 As ColorFormat, col2)
    If TypeOf col2 Is ColorFormat Then
        col1.Brightness = col2.Brightness
    Else
        col1.Brightness = col2("Brightness")
    End If
End Sub

'
' Similar to assignColorBrightness (and serving the same purpose),
' but adapted to add a Brightness object to a Collection.
'
Private Sub addBrightnessToCollection(cf As ColorFormat, coll As Collection)
    coll.Add cf.Brightness, "Brightness"
End Sub

'
' This subroutine assigns the color in the ColorFormat object
' col2 to the ColorFormat object col1. To work around issues in
' other code fragments, col2 is also accepted in the form of a
' Collection containing the same attributes as a ColorFormat
' object, but in the form of Collection items.
' Since color assignments may involve several object types (e.g.,
' shapes, text, color change effects), care must be taken in
' that the color may be specified as an index referring to the
' slide color scheme or as an RGB value. There are cases (e.g.,
' target color in color change effects) in which directly
' assigning the RGB value when the target color is indeed
' derived from the slide color scheme can result in a
' "ColorFormat: Invalid request. This object has no associated
' color scheme" runtime error, actually meaning the RGB value
' is inaccessible.
'
Private Sub assignColor(col1 As ColorFormat, ByVal col2)
    If TypeOf col2 Is ColorFormat Then
        If col2.Type = msoColorTypeRGB Then
            col1.RGB = col2.RGB
        Else
            ' I must protect from invalid assignments of color
            ' scheme indexes.
            On Error Resume Next
            col1.SchemeColor = col2.SchemeColor
            ' Sometimes the scheme color may be associated with a Brightness
            ' attribute, referred to the color in the first row of the color
            ' palette (that is, the scheme color).
            ' Apparently, this attribute is only supported starting
            ' from Office 2010 (14.0) and, in addition, may sometimes be
            ' missing altogether. There are no ways I am aware of to determine
            ' whether the color selected by the user really has this attribute,
            ' therefore the following code frament is still within the
            ' "On Error Resume Next" block.
            If Int(Mid$(Application.Version, 1, Len(Application.Version) - 2)) > 12 Then
                ' The brightness level is
                assignColorBrightness col1, col2
            End If
            On Error GoTo 0
        End If
    Else
        If col2("Type") = msoColorTypeRGB Then
            col1.RGB = col2("RGB")
        Else
            ' I must protect from invalid assignments of color
            ' scheme indexes.
            On Error Resume Next
            col1.SchemeColor = col2("SchemeColor")
            ' Sometimes the scheme color may be associated with a Brightness
            ' attribute, referred to the color in the first row of the color
            ' palette (that is, the scheme color).
            ' Apparently, this attribute is only supported starting
            ' from Office 2010 (14.0) and, in addition, may sometimes be
            ' missing altogether. There are no ways I am aware of to determine
            ' whether the color selected by the user really has this attribute,
            ' therefore the following code frament is still within the
            ' "On Error Resume Next" block.
            If Int(Mid$(Application.Version, 1, Len(Application.Version) - 2)) > 12 Then
                assignColorBrightness col1, col2
            End If
            On Error GoTo 0
        End If
    End If
            
End Sub

'
' This subroutine converts a color value from the RGB space to the
' HSL space. The result will be put in the last 3 arguments.
' The procedure is taken from http://en.wikipedia.org/wiki/HSL_and_HSV#Conversion_from_RGB_to_HSL_overview
'
Private Sub RGBtoHSL(r, g, b, h, s, l)
    max = 0: min = 255
    r = r / 255: g = g / 255: b = b / 255
    If r > max Then max = r
    If g > max Then max = g
    If b > max Then max = b
    If r < min Then min = r
    If g < min Then min = g
    If b < min Then min = b
    If max = min Then
        h = 0
    ElseIf max = r Then
        h = (60 * (g - b) / (max - min) + 360) Mod 360
    ElseIf max = g Then
        h = 60 * (b - r) / (max - min) + 120
    ElseIf max = b Then
        h = 60 * (r - g) / (max - min) + 240
    End If
    l = (max + min) / 2
    If max = min Then
        s = 0
    ElseIf l <= 1 / 2 Then
        s = (max - min) / (2 * l)
    ElseIf l > 1 / 2 Then
        s = (max - min) / (2 - 2 * l)
    End If
End Sub

'
' This subroutine converts a color value from the HSL space to the
' RGB space. The result will be put in the last 3 arguments.
' The procedure is taken from http://en.wikipedia.org/wiki/HSL_and_HSV#Conversion_from_RGB_to_HSL_overview
'
Private Sub HSLtoRGB(h, s, l, r, g, b)
    If l < 1 / 2 Then
        q = l * (1 + s)
    Else
        q = l + s - l * s
    End If
    p = 2 * l - q
    hk = h / 360
    tr = hk + 1 / 3
    ' Cannot use the Mod operator here, as it only supports integer arithmetic
    If tr < 0 Then tr = tr + 1
    If tr > 1 Then tr = tr - 1
    tg = hk
    If tg < 0 Then tg = tg + 1
    If tg > 1 Then tg = tg - 1
    tb = hk - 1 / 3
    If tb < 0 Then tb = tb + 1
    If tb > 1 Then tb = tb - 1

    If tr < 1 / 6 Then
        r = p + ((q - p) * 6 * tr)
    ElseIf tr >= 1 / 6 And tr < 1 / 2 Then
        r = q
    ElseIf tr >= 1 / 2 And tr < 2 / 3 Then
        r = p + ((q - p) * 6 * (2 / 3 - tr))
    Else
        r = p
    End If
    If tg < 1 / 6 Then
        g = p + ((q - p) * 6 * tg)
    ElseIf tg >= 1 / 6 And tg < 1 / 2 Then
        g = q
    ElseIf tg >= 1 / 2 And tg < 2 / 3 Then
        g = p + ((q - p) * 6 * (2 / 3 - tg))
    Else
        g = p
    End If
    If tb < 1 / 6 Then
        b = p + ((q - p) * 6 * tb)
    ElseIf tb >= 1 / 6 And tb < 1 / 2 Then
        b = q
    ElseIf tb >= 1 / 2 And tb < 2 / 3 Then
        b = p + ((q - p) * 6 * (2 / 3 - tb))
    Else
        b = p
    End If
    r = r * 255: g = g * 255: b = b * 255
End Sub

'
' This subroutine converts a color value represented by VBA as a Long
' integer into its RGB components. The result is put in the last
' 3 arguments of the subroutine.
'
Private Sub colToRGB(col, r, g, b)
    r = col Mod 256
    g = (col \ 256) Mod 256
    b = (col \ 256 \ 256) Mod 256
End Sub

'
' This subroutine "rotates" the hue of a given color of the
' specified angle (in degrees).
'
Private Sub rotateColor(col As ColorFormat, rot)
    colToRGB col.RGB, r, g, b
    RGBtoHSL r, g, b, h, s, l
    h = (h + rot) Mod 360
    HSLtoRGB h, s, l, r, g, b
    col.RGB = RGB(r, g, b)
End Sub

'
' This subroutine alters the lightness of a given color.
' The amount should be between 0 and 1.
'
Private Sub changeLightness(col As ColorFormat, amount)
    colToRGB col.RGB, r, g, b
    RGBtoHSL r, g, b, h, s, l
    l = l + amount
    If l > 1 Then l = 1
    If l < 0 Then l = 0
    HSLtoRGB h, s, l, r, g, b
    col.RGB = RGB(r, g, b)
End Sub

'
' After a motion effect has been applied to a shape, the coordinates
' of all subsequent motion effects have been moved together with the
' shape. This subroutine applies a given shift to the arrival
' coordinates (indeed, arrival coordinates is all I need to update)
' of all the other motion effects for the same shape. Arguments
' effectSequence (the sequence of effects applied to the shape) and
' sh (the affected shape) do not need, and in general do not, refer
' to the same slide.
'
' A motion path is specified in VML. Information about the specification
' can be found here: http://www.w3.org/TR/NOTE-VML#_Toc416858391
'
Private Sub shiftAllMotions(effectSequence As Sequence, sh As Shape, shiftX, shiftY)
    Dim currentEffect As effect, lastX As Double, lastY As Double
    For Each currentEffect In effectSequence
        ' The following variable is where I will put the reconstructed
        ' path with updated arrival coordinates
        motionPathString$ = ""
        ' Keep in mind that sh is a shape the effect is applied to (therefore
        ' it comes from a certain slide), while effectSequence is the sequence of effects
        ' under consideration (which comes from a different slide). Therefore,
        ' operator "Is" cannot be used to match the shapes whose motion effects
        ' should be updated.
        If isPathEffect(currentEffect) And currentEffect.Shape.Id = sh.Id Then
            ' This is a motion effect applied to the shape under consideration
            motionPathTokens = Split(currentEffect.Behaviors(1).MotionEffect.Path)
            ' The first character states this is a path motion, therefore I preserve it
            motionPathString$ = motionPathString$ + Trim(motionPathTokens(0)) + " "
            If currentEffect.Behaviors(1).Timing.Speed < 0 Then
                ' The path has been reversed: update origin coordinates instead
                lastX = localizeDecimalSeparators(motionPathTokens(1))
                lastY = localizeDecimalSeparators(motionPathTokens(2))
                lastX = lastX + shiftX
                lastY = lastY + shiftY
                motionPathString$ = motionPathString$ + Trim(Str$(lastX)) + " " + Trim(Str$(lastY)) + " "
                ' Append the rest of the motion string
                For i = 3 To UBound(motionPathTokens)
                    motionPathString$ = motionPathString$ + motionPathTokens(i) + " "
                Next i
            Else
                ' Update the last two (i.e., arrival) coordinates
                getLastCoordinates currentEffect.Behaviors(1).MotionEffect.Path, lastX, lastY, lastToken
                lastX = lastX + shiftX
                lastY = lastY + shiftY
                ' Copy everything but the last two coordinates from the original
                ' motion string
                For i = 0 To lastToken
                    motionPathString$ = motionPathString$ + motionPathTokens(i) + " "
                Next i
                ' Append the modified coordinates
                motionPathString$ = motionPathString$ + Trim(Str$(lastX)) + " " + Trim(Str$(lastY)) + " "
            End If
            ' Assign the new path
            currentEffect.Behaviors(1).MotionEffect.Path = motionPathString$
        End If
    Next currentEffect
End Sub

'
' This converts an angle from degrees to radians. At the
' same time, since shape rotation angles are computed in PowerPoint
' starting from the positive Y semiaxis and going in
' clockwise direction, it reverses the convention by returning
' an angle in radiants that starts from the positive X semiaxis
' and goes counterclockwise.
'
Private Function degToRad(degAngle) As Double
    degToRad = 3.14159265358979 * ((360 - degAngle) Mod 360) / 180
End Function

'
' This subroutine gets the last (i.e., arrival) coordinates from
' a string describing a motion path. Extracted coordinates are put
' in lastX and lastY, while lastTokenBeforeCoordinates will be
' updated with the index of the token in pathString$ that precedes
' the last coordinates.
'
Private Sub getLastCoordinates(pathString$, lastX As Double, lastY As Double, lastTokenBeforeCoordinates)
    pathStringTokens = Split(pathString$)
    tokenIndex = UBound(pathStringTokens)
    While tokenIndex > 0
        If pathStringTokens(tokenIndex) <> "" And _
            Not (Mid$(pathStringTokens(tokenIndex), 1, 1) >= "A" And _
            Mid$(pathStringTokens(tokenIndex), 1, 1) <= "Z") Then
            lastY = localizeDecimalSeparators(pathStringTokens(tokenIndex))
            lastX = localizeDecimalSeparators(pathStringTokens(tokenIndex - 1))
            lastTokenBeforeCoordinates = tokenIndex - 2
            Exit Sub
        End If
        tokenIndex = tokenIndex - 1
    Wend
End Sub


'
' This subroutine does what it says: it applies an emphasis
' (or motion) effect to a shape. Arguments are:
' - seq: a sequence of effects which will only be modified to update
'   motion path coordinates for the specific case of a motion effect
' - e: the emphasis effect to be applied
' - sh: the shape it applies to
'
Private Sub applyEmphasisEffect(seq As Sequence, e As effect, sh As Shape, final_colors As Collection)
    On Error GoTo recover
    ePar = getEffectParagraph(e)
    ' Here I should be supposed to check the value of
    ' e.Shape.HasTextFrame before attemping to access
    ' the sh.TextFrame.TextRange property. Unfortunately,
    ' in some cases PowerPoint returns false even if
    ' properties like sh.TextFrame.TextRange.Font.Size
    ' can be accessed.
    ' Worked around by attempting assignments anyway, and
    ' watching for errors during the process.
    On Error Resume Next
    shTextRange = Null
    shTextRange2 = Null
    If ePar > 0 Then
        ' This effect applies to a text paragraph
        Set shTextRange = sh.TextFrame.TextRange.Paragraphs(ePar)
        Set shTextRange2 = sh.TextFrame2.TextRange.Paragraphs(ePar)
    Else
        Set shTextRange = sh.TextFrame.TextRange
        Set shTextRange2 = sh.TextFrame2.TextRange
    End If
    On Error GoTo recover
    ' Note: if an effect acts both on a text element and on its container
    ' shape, then the effect must first be applied to the container shape,
    ' in order to avoid unpredictable automatic resizing.
    If e.EffectType = msoAnimEffectGrowShrink Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            ' I am not scaling a bitmap here, therefore I need to
            ' recompute map X and Y scaling in accordance with the shape
            ' rotation.
            rotCos = Cos(degToRad(sh.Rotation))
            rotSin = Sin(degToRad(sh.Rotation))
            scaleX = e.Behaviors(1).ScaleEffect.ByX / 100 * Abs(rotCos) + e.Behaviors(1).ScaleEffect.ByY / 100 * Abs(rotSin)
            scaleY = e.Behaviors(1).ScaleEffect.ByX / 100 * Abs(rotSin) + e.Behaviors(1).ScaleEffect.ByY / 100 * Abs(rotCos)
            ' Disable size autofitting for text frames and unlock
            ' aspect ratio
            sh.LockAspectRatio = msoFalse
            On Error Resume Next
            sh.TextFrame.AutoSize = ppAutoSizeNone
            On Error GoTo recover
            sh.ScaleWidth scaleX, msoFalse, msoScaleFromMiddle
            sh.ScaleHeight scaleY, msoFalse, msoScaleFromMiddle
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Size = shTextRange.Font.Size * (e.Behaviors(1).ScaleEffect.ByX / 100 + e.Behaviors(1).ScaleEffect.ByY / 100) / 2
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectChangeFontColor Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                assignColor shTextRange.Font.Color, final_colors
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectChangeFillColor Then
        If sh.Fill.Transparency < 1 Then
            sh.Fill.Solid
        End If
        ' Use the final_colors Collection here, to retrieve the
        ' correct target color value
        assignColor sh.Fill.ForeColor, final_colors
        ' Original statement follows
'        assignColor sh.Fill.ForeColor, e.EffectParameters.Color2
    ElseIf e.EffectType = msoAnimEffectChangeFontStyle Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Italic = (e.Behaviors(1).SetEffect.To = 1)
                shTextRange.Font.Bold = (e.Behaviors(2).SetEffect.To = 1)
                shTextRange.Font.Underline = (e.Behaviors(3).SetEffect.To = 1)
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectTransparency Then
        ' Potentially bad consequence: objects that are made totally
        ' transparent cannot have their transparency changed
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Line.Transparency < 1 Then
                sh.Line.Transparency = e.EffectParameters.amount
            End If
            If sh.Fill.Transparency < 1 Then
                sh.Fill.Transparency = e.EffectParameters.amount
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange2 = Null
            On Error Resume Next
            Set shTextRange2 = sh.GroupItems(shapeID).TextFrame2.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange2) Then
                If shTextRange2.Font.Fill.Transparency < 1 Then
                    shTextRange2.Font.Fill.Transparency = e.EffectParameters.amount
                End If
                If shTextRange2.Font.Line.Transparency < 1 Then
                    shTextRange2.Font.Line.Transparency = e.EffectParameters.amount
                End If
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange2 = Null
                    On Error Resume Next
                    Set shTextRange2 = sh.GroupItems(shapeID).TextFrame2.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectChangeFont Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Name = e.EffectParameters.FontName
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectChangeLineColor Then
        If Not sh.Line.Visible Then sh.Line.Visible = msoTrue
        ' Use the final_colors Collection here, to retrieve the
        ' correct target color value
        assignColor sh.Line.ForeColor, final_colors
        ' Original statement follows
'        assignColor sh.Line.ForeColor, e.EffectParameters.Color2
    ElseIf e.EffectType = msoAnimEffectChangeFontSize Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                ' Please leave the /1 alone: it is required for some strange internal
                ' type conversion, otherwise leading to improper font sizes :-(
                shTextRange.Font.Size = shTextRange.Font.Size * e.Behaviors(1).PropertyEffect.To / 1
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectSpin Then
        ' Rotating just the text is not supported
        sh.Rotation = sh.Rotation + e.Behaviors(1).RotationEffect.By
    ElseIf e.EffectType = msoAnimEffectDesaturate Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                With sh.Fill.ForeColor
                    colToRGB .RGB, r, g, b
                    .RGB = RGB((r + g + b) / 3, (r + g + b) / 3, (r + g + b) / 3)
                End With
                With sh.Fill.BackColor
                    colToRGB .RGB, r, g, b
                    .RGB = RGB((r + g + b) / 3, (r + g + b) / 3, (r + g + b) / 3)
                End With
            End If
            If sh.Line.Transparency < 1 Then
                With sh.Line.ForeColor
                    colToRGB .RGB, r, g, b
                    .RGB = RGB((r + g + b) / 3, (r + g + b) / 3, (r + g + b) / 3)
                End With
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                With shTextRange.Font.Color
                    colToRGB .RGB, r, g, b
                    .RGB = RGB((r + g + b) / 3, (r + g + b) / 3, (r + g + b) / 3)
                End With
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectColorWave Or e.EffectType = msoAnimEffectColorBlend Or _
            e.EffectType = msoAnimEffectBrushOnColor Or e.EffectType = msoAnimEffectTeeter Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                assignColor sh.Fill.ForeColor, e.EffectParameters.Color2
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                assignColor shTextRange.Font.Color, final_colors
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectComplementaryColor2 Then
        ' PowerPoint computes the complementary color in some other way.
        ' I feel pretty satisfied with this rotation in the HSL space
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                rotateColor sh.Fill.ForeColor, 180
            End If
            If sh.Line.Transparency < 1 Then
                rotateColor sh.Line.ForeColor, 180
            End If
        End If
    ElseIf e.EffectType = msoAnimEffectVerticalGrow Then
        ' Font scaling alone is not supported for this effect
        
        ' Disable size autofitting for text frames and unlock
        ' aspect ratio
        sh.LockAspectRatio = msoFalse
        On Error Resume Next
        sh.TextFrame.AutoSize = ppAutoSizeNone
        On Error GoTo recover
        sh.ScaleHeight 1.5, msoFalse
        shiftY = sh.Height / 4
        If sh.Fill.Transparency < 1 Then
            assignColor sh.Fill.ForeColor, e.EffectParameters.Color2
        End If
        sh.Top = sh.Top - shiftY
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                assignColor shTextRange.Font.Color, final_colors
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectLighten Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                changeLightness sh.Fill.ForeColor, 0.3
            End If
            If sh.Line.Transparency < 1 Then
                changeLightness sh.Line.ForeColor, 0.3
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                changeLightness shTextRange.Font.Color, 0.3
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectBrushOnUnderline Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Underline = msoTrue
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectComplementaryColor Then
        ' PowerPoint computes the complementary color in some other way.
        ' I feel pretty satisfied with this rotation in the HSL space
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                rotateColor sh.Fill.ForeColor, 120
            End If
            If sh.Line.Transparency < 1 Then
                rotateColor sh.Line.ForeColor, 120
            End If
        End If
    ElseIf e.EffectType = msoAnimEffectContrastingColor Then
        ' PowerPoint computes the contrasting color in some other way.
        ' I feel pretty satisfied with this rotation in the HSL space
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                rotateColor sh.Fill.ForeColor, 90
            End If
            If sh.Line.Transparency < 1 Then
                rotateColor sh.Line.ForeColor, 90
            End If
        End If
    ElseIf e.EffectType = msoAnimEffectBoldFlash Then
        ' msoAnimEffectBoldFlash is a non-permanent effect
    ElseIf e.EffectType = msoAnimEffectFlashBulb Then
        ' msoAnimEffectFlashBulb is a non-permanent effect
    ElseIf e.EffectType = msoAnimEffectDarken Then
        If e.Shape.Type = msoPlaceholder Or e.EffectInformation.AnimateBackground Or Not e.Shape.TextFrame.HasText Or e.Shape.Type = msoGroup Then
            If sh.Fill.Transparency < 1 Then
                changeLightness sh.Fill.ForeColor, -0.3
            End If
            If sh.Line.Transparency < 1 Then
                changeLightness sh.Line.ForeColor, -0.3
            End If
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                changeLightness shTextRange.Font.Color, -0.3
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectGrowWithColor Then
        If sh.Fill.Transparency < 1 Then
            sh.Fill.Solid
            assignColor sh.Fill.ForeColor, e.EffectParameters.Color2
        End If
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Size = shTextRange.Font.Size * 1.5
                assignColor shTextRange.Font.Color, final_colors
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectFlicker Then
        ' msoAnimEffectFlicker is a non-permanent effect
    ' *** WARNING: the shaking effect has no associated effecttype (PowerPoint bug :-((( )
    ElseIf e.EffectType = msoAnimEffectBoldReveal Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Bold = msoTrue
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ElseIf e.EffectType = msoAnimEffectWave Then
        ' msoAnimEffectWave is a non-permanent effect
    ElseIf e.EffectType = msoAnimEffectStyleEmphasis Then
        ' Font effects may be applied to a group. In that case,
        ' at least for versions of PowerPoint prior to 2007, we
        ' are forced to apply the effect for each member of the
        ' group.
        shapeID = 1
        If sh.Type = msoGroup Then
            shTextRange = Null
            On Error Resume Next
            Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
            On Error GoTo recover
        End If
        Do
            If Not IsNull(shTextRange) Then
                shTextRange.Font.Italic = msoTrue
                shTextRange.Font.Bold = msoTrue
                shTextRange.Font.Underline = msoTrue
                assignColor shTextRange.Font.Color, e.EffectParameters.Color2
            End If
            shapeID = shapeID + 1
            If sh.Type = msoGroup Then
                If shapeID > sh.GroupItems.Count Then
                    shapeID = 0
                Else
                    shTextRange = Null
                    On Error Resume Next
                    Set shTextRange = sh.GroupItems(shapeID).TextFrame.TextRange
                    On Error GoTo recover
                End If
            Else
                shapeID = 0
            End If
        Loop Until shapeID = 0
    ' *** WARNING: the blinking effect has no associated effecttype (PowerPoint bug :-((( )
    ElseIf e.EffectType = msoAnimEffectBlast Then
        ' msoAnimEffectBlast has too vague a behavior to be implemented :-O
    Else
        If isEmphasisEffect(e) Then
            On Error GoTo 0
            ' Ok, this is neither an emphasis effect nor an entry effect:
            ' it must be a motion effect
            motionpath = Split(e.Behaviors(1).MotionEffect.Path)
            Dim lastX As Double, lastY As Double
            If e.Behaviors(1).Timing.Speed < 0 Then
                lastX = localizeDecimalSeparators(motionpath(1))
                lastY = localizeDecimalSeparators(motionpath(2))
            Else
                getLastCoordinates e.Behaviors(1).MotionEffect.Path, lastX, lastY, lastToken
            End If
            ' Coordinates are expressed in VML (see http://www.w3.org/TR/1998/NOTE-VML-19980513#_Toc416858391)
            ' as multiples of the slide width/height and are relative to the shape center
            shapeCenterX = (sh.Left + sh.Width / 2) / ActivePresentation.PageSetup.SlideWidth
            shapeCenterY = (sh.Top + sh.Height / 2) / ActivePresentation.PageSetup.SlideHeight
            newX = (shapeCenterX + lastX) * ActivePresentation.PageSetup.SlideWidth
            newY = (shapeCenterY + lastY) * ActivePresentation.PageSetup.SlideHeight
            sh.Left = newX - sh.Width / 2
            sh.Top = newY - sh.Height / 2
            shiftAllMotions seq, sh, -lastX, -lastY
        End If
    End If
    Exit Sub
recover:
    ' Ok, Powerpoint bug again: this is an emphasis effect that
    ' has no EffectType member. Let's pass it by.
End Sub

'
' This function returns true if (and only if) the effect given
' as argument is a motion (path) effect
'
Private Function isPathEffect(e As effect) As Boolean
    On Error GoTo pathRecover
    isPathEffect = False
    ' The following conditions have been built starting from the page "Powerpoint
    ' constants" of the VBA documentation.
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPath5PointStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCrescentMoon
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSquare
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTrapezoid
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathHeart
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathOctagon
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPath6PointStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathFootball
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathEqualTriangle
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathParallelogram
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathPentagon
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPath4PointStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPath8PointStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTeardrop
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathPointyStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvedSquare
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvedX
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathVerticalFigure8
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvyStar
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathLoopdeLoop
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathBuzzsaw
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathHorizontalFigure8
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathPeanut
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathFigure8Four
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathNeutron
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSwoosh
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathBean
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathPlus
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathInvertedTriangle
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathInvertedSquare
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTurnRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathArcDown
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathZigzag
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSCurve2
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSineWave
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathBounceLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathDown
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTurnUp
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathArcUp
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathHeartbeat
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSpiralRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathWave
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvyLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathDiagonalDownRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTurnDown
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathArcLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathFunnel
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSpring
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathBounceRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSpiralLeft
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathDiagonalUpRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathTurnUpRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathArcRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathSCurve1
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathDecayingWave
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathCurvyRight
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathStairsDown
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathUp
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectPathRight

    ' 0 = msoAnimEffectCustom = Customized path
    isPathEffect = isPathEffect Or e.EffectType = msoAnimEffectCustom
    Exit Function
    
pathRecover:
    ' Powerpoint bug: this effect has no EffectType property;
    ' I cannot either recognize or handle it. At the time of
    ' writing this code, there were no motion effects affected
    ' by this problem, therefore this is not a motion effect.
    isPathEffect = False
End Function

'
' This function returns true iff the given effect is either
' an emphasis effect or a motion effect.
'
Private Function isEmphasisEffect(e As effect) As Boolean
    On Error GoTo recoverIsEmphasis
    isEmphasisEffect = False
    ' The following conditions have been built starting from the page "Powerpoint
    ' constants" of the VBA documentation.
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectGrowShrink
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFontColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFillColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFontStyle
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectTransparency
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFont
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeLineColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectChangeFontSize
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectSpin
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectDesaturate
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectColorWave
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectComplementaryColor2
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectVerticalGrow
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectLighten
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectColorBlend
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBrushOnUnderline
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBrushOnColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectComplementaryColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectContrastingColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBoldFlash
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectFlashBulb
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectDarken
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectGrowWithColor
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectTeeter
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectFlicker
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBoldReveal
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectWave
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectStyleEmphasis
    isEmphasisEffect = isEmphasisEffect Or e.EffectType = msoAnimEffectBlast
    
    isEmphasisEffect = isEmphasisEffect Or isPathEffect(e)

    ' If isEmphasisEffect is true at this point, then I have
    ' an emphasis or motion effect. But let's really make sure it is not
    ' an entry/exit effect.
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectAppear
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFly
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectBlinds
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectBox
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCheckerboard
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCircle
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCrawl
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectDiamond
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectDissolve
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFade
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFlashOnce
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectPeek
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectPlus
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectRandomBars
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSpiral
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSplit
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectStretch
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectStrips
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSwivel
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectWedge
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectWheel
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectWipe
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectZoom
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectRandomEffects
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectBoomerang
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectBounce
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectColorReveal
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCredits
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectEaseIn
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFloat
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectGrowAndTurn
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectLightSpeed
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectPinwheel
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectRiseUp
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSwish
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectThinLine
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectUnfold
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectWhip
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectAscend
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectCenterRevolve
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFadedSwivel
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectDescend
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSling
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectSpinner
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectStretchy
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectZip
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectArcUp
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFadedZoom
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectGlide
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectExpand
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFlip
    isEmphasisEffect = isEmphasisEffect And e.EffectType <> msoAnimEffectFold
    Exit Function
recoverIsEmphasis:
    ' Powerpoint bug: this effect has no EffectType property;
    ' I cannot either recognize or handle it. Luckily enough,
    ' there is no need to process the affected effects because
    ' they are non-permanent (apart from the color that the
    ' shaking effect allows to apply to the shape). Here I
    ' assume that an unrecognizable effect is an emphasis effect.
    isEmphasisEffect = True
End Function

'
' This function takes an effect as argument. If the
' effect is applied to a text paragraph, it returns the
' index of that text paragraph (in its container shape).
' Otherwise, it returns -1.
'
Private Function getEffectParagraph(e As effect)
    paragraph_idx = -1
    On Error Resume Next
    ' The following assignment may fail because the Paragraph property does not
    ' exist at all for those effects that are applied to shapes instead of text.
    ' But, was this truly expected by design? :-?
    paragraph_idx = e.Paragraph
    On Error GoTo 0
    getEffectParagraph = paragraph_idx
End Function

'
' Utility function that returns the ID of the shape to which an effect
' is applied, followed by a comma, followed by the paragraph number to
' which the effect is applied (or 0 if the effect applies to the whole
' shape).
Public Function getFullShapeID(e As effect)
    Dim paragraph_number As Integer
    Dim fullShapeID As String
    paragraph_number = getEffectParagraph(e)
    fullShapeID = LTrim(Str$(e.Shape.Id)) + ","
    If paragraph_number < 0 Then
        fullShapeID = fullShapeID + "0"
    Else
        fullShapeID = fullShapeID + LTrim(Str$(paragraph_number))
    End If
    getFullShapeID = fullShapeID
End Function

'
' This subroutine removes all the animation effects from a slide. Useful
' to leave slides clean after processing
'
Private Sub purgeEffects(s As slide)
    For i = 1 To s.timeline.MainSequence.Count
        s.timeline.MainSequence(1).Delete
    Next i
    s.SlideShowTransition.EntryEffect = ppEffectNone
End Sub

'
' This function moves elements from slide masters to slides, in order to keep slide
' numbers fixed during the split. Note that slide numbers may occur in several shapes
' in a slide master, not just the "slide number" footer: slide numbers appearing in
' such extra shapes will not be processed.
'
Private Sub bakeSlideNumbers(slide_range As SlideRange)
    Dim sh As Shape

    ProgressForm.infoLabel = "Adjusting slide numbers. This may take some time..."
    DoEvents

    ' Placeholders from slide masters (even custom layouts) appear as standard shapes in the
    ' slides. Therefore, here we search for placeholder shapes in each slide and, when found,
    ' we simply reassign the text to the shape, in order to turn any special <pagenumber>
    ' field into plain text
    processed_slides = 0
    For Each s In slide_range
        For Each sh In s.Shapes
            If sh.Type = msoPlaceholder Then
                With sh.PlaceholderFormat
                    If .Type = ppPlaceholderSlideNumber Or _
                       .Type = ppPlaceholderDate Or _
                       .Type = ppPlaceholderFooter Then
                       ' Text is baked character by character, in order to avoid losing formatting
                       For c = 1 To sh.TextFrame.TextRange.Characters.Count
                            sh.TextFrame.TextRange.Characters(c) = sh.TextFrame.TextRange.Characters(c)
                       Next c
                    End If
                End With
            End If
        Next sh
        processed_slides = processed_slides + 1
        DoEvents
    Next s

    ProgressForm.infoLabel = ""
    DoEvents
End Sub

'
' This function enriches existing slide numbers with a subindex, namely a progressive
' number assigned anew to each slide resulting from splitting a single original one.
' It works in close conjunction with bakeSlideNumbers, with a main difference:
' - bakeSlideNumbers is invoked once on all the slide deck to make slide numbers
'   persistent
' - augmentSlideNumbers is invoked once for each split slide, strictly after processing
'   of that slide has finished and, possibly, after a duplicate of that slide is
'   generated (modified slide numbers would otherwise be inherited in all subsequent
'   slides)
'
Private Sub augmentSlideNumbers(current_slide As slide, progressive_slide_count)
    Dim sh As Shape

    For Each sh In current_slide.Shapes
        If sh.Type = msoPlaceholder Then
            With sh.PlaceholderFormat
                If slideNumbersAdjustMode = SLIDENUMBER_SUBINDEX And .Type = ppPlaceholderSlideNumber Then
                    sh.TextFrame.TextRange.InsertAfter "." + Right$(Str$(progressive_slide_count), Len(Str$(progressive_slide_count)) - 1)
                End If
            End With
        End If
    Next sh
End Sub

' Delete all shapes that are supposed to be appear later than current_effect
' in the animation timeline or that have already disappeared by current_effect.
' Deletions are applied to shapes contained in target_slide
Private Sub purgeInvisibleShapes(ByRef shape_visible As Collection, timeline As Sequence, target_slide As slide)
    Dim e As effect
    Dim par As Integer
    Dim target_shape As Shape
    
    ' Iterating on each shape to which an animation effect is applied
    ' would be enough here. Yet, since there are no ways to retrieve
    ' the keys in a collection, here we iterate on each effect instead
    For Each e In timeline
        If Not shape_visible(getFullShapeID(e)) Then
            Set target_shape = findShapeByIDinTag(e.Shape.Id, target_slide)
            ' Each shape may appear multiple times in the animation timeline,
            ' therefore it might have already been deleted at a previous
            ' iteration.
            ' Note that processing each shape more than once may not
            ' necessarily be redundant, as different animation steps may
            ' for example affect different paragraphs of the same shape.
            If Not target_shape Is Nothing Then
                par = getEffectParagraph(e)
                If par > 0 Then
                    ' Completely removing the paragraph (e.g., by clearing its text or
                    ' making it invisible) is not correct, since it must still
                    ' take up space to let the rest of the text in the frame stay
                    ' where it currently is
                    target_shape.TextFrame2.TextRange.Paragraphs(par).Font.Fill.Transparency = 1
                    ' Sometimes, especially when (small) images are used, bullets
                    ' stay visible even after executing the above statement. Therefore,
                    ' here we try again to hide them. Once more, removing them
                    ' (i.e., clearing them or making them invisible) is not a good idea,
                    ' because it would cause the corresponding paragraph text to shift
                    ' leftwards and it would cause numbering in a list to be mixed up.
                    target_shape.TextFrame2.TextRange.Paragraphs(par).ParagraphFormat.Bullet.Font.Size = 1
                Else
                    target_shape.Delete
                End If
            End If
        End If
    Next e
End Sub

'
' Returns true if an effect is mouse-triggered
'
Private Function isMouseTriggered(effect As effect)
    With effect.Timing
        isMouseTriggered = .TriggerType <> msoAnimTriggerAfterPrevious And _
                           .TriggerType <> msoAnimTriggerWithPrevious
    End With
End Function

'
' Pre-process effects in timelines in order to:
' - remove non-persistent effects (i.e., rewound after playing, autoreverse)
' - add an extra "fake" exit animation for those effects that are
'   set to "hide on next click"
' - turn entry effects for which the "hide after playing" property
'   is set into exit effects
'
Private Sub preprocessEffects(current_slide As slide, final_colors As Collection)
    Dim current_effect As effect, e As effect, e2 As effect, insert_after_effect As effect
    
    ' Iterate over all the effects by using an index. Using a native iterator is not
    ' possible here because the list of effects is manipulated during the iteration
    ' itself.
    effects_count = current_slide.timeline.MainSequence.Count
    i = 1
    While i <= effects_count
        Set current_effect = current_slide.timeline.MainSequence(i)
        
        If current_effect.EffectInformation.AfterEffect = msoAnimAfterEffectHideOnNextClick Then
            ' Whatever its native type, this effect is set to hide the
            ' affected shape after the next mouse click. In order to
            ' be able to process this later, a "fake" extra exit effect is
            ' added after the next mouse-triggered effect
            Set insert_after_effect = Nothing
            ' Only consider effects that follow the current one
            For i2 = current_slide.timeline.MainSequence.Count To current_effect.Index + 1 Step -1
                If isMouseTriggered(current_slide.timeline.MainSequence(i2)) Then
                    Set insert_after_effect = current_slide.timeline.MainSequence(i2)
                End If
            Next i2
            If insert_after_effect Is Nothing Then
                ' There were no more mouse-triggered effects before the
                ' end of the timeline, therefore the "fake" effect is added
                ' at the end of the timeline, but only played at a mouse click
                Set insert_after_effect = current_slide.timeline.MainSequence(current_slide.timeline.MainSequence.Count)
                Set e2 = current_slide.timeline.MainSequence.AddEffect(current_effect.Shape, msoAnimEffectDissolve, , msoAnimTriggerOnPageClick)
            Else
                Set e2 = current_slide.timeline.MainSequence.AddEffect(current_effect.Shape, msoAnimEffectDissolve, , msoAnimTriggerWithPrevious)
            End If
            effects_count = effects_count + 1
            e2.Exit = msoTrue
            ' Best thing would be to insert the exit effect right after the next click-triggered
            ' effect, but apparently the Index argument of AddEffect may be handled unpredictably.
            ' So, we need to work this around by inserting the effect at the end of the sequence and,
            ' only afterwards, move it to the right location.
            e2.MoveAfter insert_after_effect
            
            ' Since indexes of animation effects have been updated as a consequence of the
            ' effect insertion, the Collection that stores final colors for emphasis effects
            ' has to be updated accordingly
            With final_colors(Str$(current_slide.SlideID))
                For e_idx = current_slide.timeline.MainSequence.Count To insert_after_effect.Index + 2 Step -1
                    .Add .Item(Str$(e_idx - 1)), Str$(e_idx)
                    .Remove Str$(e_idx - 1)
                Next e_idx
                ' Avoid leaving an empty slot
                .Add Nothing, Str$(insert_after_effect.Index + 1)
            End With
        End If
        
        If current_effect.EffectInformation.AfterEffect = msoAnimAfterEffectHide Then
            ' This effect behaves as an exit effect: unless it is already
            ' an exit effect, replace it with an exit effect
            If Not current_effect.Exit Then
                Set e2 = current_slide.timeline.MainSequence.AddEffect(current_effect.Shape, msoAnimEffectDissolve, , msoAnimTriggerWithPrevious)
                e2.MoveAfter current_effect
                current_effect.Delete
                
                ' No updates of the final_colors Collection are due, since an
                ' effect replacement has taken place (for which the Color2 property
                ' will not even be taken into account)
            End If
        End If

        ' Rewound-at-end and autoreversed (emphasis) effects have no persistent impact on the shape
        ' they are applied to, unless they are also set, e.g., for hiding on next click (which has
        ' already been checked). Their presence is still required, though, as they may determine
        ' timeline advancement steps by mouse clicks. Therefore, such an effect is replaced with
        ' another one that simply has no persistent effects (but will still be processed in the
        ' following).
        If current_effect.Timing.AutoReverse Or current_effect.Timing.RewindAtEnd Then
            current_effect.EffectType = msoAnimEffectFlashBulb
        End If
        
        i = i + 1
    Wend
End Sub

'
' Store Color2 attributes of all emphasis effects in the slide deck
' to a separate Collection object. This is required to work around an
' issue for which Color2 properties are corrupted when a reference
' to a Slide object is defined (which happens quite often in the rest
' of the code). For this reason, Slide object attributes are always
' addressed using the full object path ActivePresentation.Slides(x)...
' below.
'
Private Sub saveAllFinalColors(final_colors As Collection)
    Set final_colors = New Collection
    Dim current_effect As effect, cf As ColorFormat, stored_cf As Collection, _
        final_colors_per_effect As Collection, final_colors_for_current_effect As Collection
    For slide_index = 1 To ActivePresentation.Slides.Count
        Set final_colors_per_effect = New Collection
        For effect_index = 1 To ActivePresentation.Slides(slide_index).timeline.MainSequence.Count
            Set cf = Nothing
            cType = -1000
            Set final_colors_for_current_effect = Nothing
            ' Attempt accessing the Color2 attribute, if the effect has one.
            ' Unfortunately, no better way seems to exist other than trying
            ' and detecting an error in the access attempt.
            On Error Resume Next
            Set cf = ActivePresentation.Slides(slide_index).timeline.MainSequence(effect_index).EffectParameters.Color2
            ' Sometimes the Color2 object may exist but its Type attribute may not
            cType = ActivePresentation.Slides(slide_index).timeline.MainSequence(effect_index).EffectParameters.Color2.Type
            On Error GoTo 0
            If Not cf Is Nothing And cType <> -1000 Then
                Set final_colors_for_current_effect = New Collection
                With final_colors_for_current_effect
                    .Add cf.Type, "Type"
                    If cf.Type = msoColorTypeRGB Then
                        .Add cf.RGB, "RGB"
                    Else
                        On Error Resume Next
                        .Add cf.SchemeColor, "SchemeColor"
                        ' Sometimes the scheme color may be associated with a Brightness
                        ' attribute, referred to the color in the first row of the color
                        ' palette (that is, the scheme color).
                        ' Apparently, this attribute is only supported starting
                        ' from Office 2010 (14.0) and, in addition, may sometimes be
                        ' missing altogether. There are no ways I am aware of to determine
                        ' whether the color selected by the user really has this attribute,
                        ' therefore the following code frament is still within the
                        ' "On Error Resume Next" block.
                        If Int(Mid$(Application.Version, 1, Len(Application.Version) - 2)) > 12 Then
                            addBrightnessToCollection cf, final_colors_for_current_effect
                        End If
                        On Error GoTo 0
                    End If
                End With
            End If
            final_colors_per_effect.Add final_colors_for_current_effect, Str$(effect_index)
        Next effect_index
        final_colors.Add final_colors_per_effect, Str$(ActivePresentation.Slides(slide_index).SlideID)
    Next slide_index
End Sub

'
' In some cases duplicating a slide results in re-generating the IDs
' of the shapes it contains. To prevent this, the following function
' preserves shape IDs into tags.
'
Private Sub copyShapeIDsToTags(current_slide As slide)
    Dim sh As Shape
    For Each sh In current_slide.Shapes
        sh.Tags.Add "ID", Str$(sh.Id)
    Next sh
End Sub

'
' Refresh progress bar representing the progress for the current slides
'
' Warning: the VBA interpreter has an issue with Double type variables on
' MacOS (see, for example, https://techcommunity.microsoft.com/t5/excel/runtime-error-6-overflow-with-dim-double-macos-catalina-excel/m-p/786433).
' As a workaround, argument types for this function are intentionally
' undeclared.
Private Sub setProgressBar(current_value, max_value)
    Dim percentage As Integer
    
    percentage = CInt(current_value / max_value * 100)
    ProgressForm.OverallLabel = Str$(percentage) + " %"
    ProgressForm.OverallBar.Width = percentage / 100 * maxProgressWidth
    DoEvents
End Sub

'
' Main loop
'
Sub PPspliT_main()

    On Error GoTo error_handler

    If Application.Presentations.Count = 0 Then
        ' No open presentations
        Exit Sub
    End If

    Dim slide_range As SlideRange
    cancelStatus = False
    
    ' Save the contents of any Color2 effect attributes in a separate
    ' data structure for later retrieval. A few tests have highlighted
    ' that setting references to a Slide object can corrupt the contents
    ' of Color2 objects for emphasis effects that make use of this
    ' attribute. For example, consider the following assignment:
    '   Set slide_range = ActiveWindow.Presentation.Slides.Range
    ' The following sample instructions have a different outcome
    ' depending on whether they are executed before or after
    ' the above assignment:
    '   Debug.Print ActivePresentation.Slides(7).timeline.MainSequence(2).EffectParameters.Color2.RGB
    '   Debug.Print ActivePresentation.Slides(7).timeline.MainSequence(4).EffectParameters.Color2.RGB
    ' If they are executed before the assignment, then two
    ' different color values are printed, as expected.
    ' If they are executed after the assignment, then the
    ' second instruction always prints the same value as the
    ' first one, regardless of the specific animation sequence
    ' index used in the two instructions. Apparently, any
    ' future attempts to access the Color2 object's properties
    ' for any shape will likely return this same value.
    ' For this reason, Color2 objects are retrieved and stored
    ' in a separate Collection for being used later on.
    Dim final_colors As Collection
    saveAllFinalColors final_colors

    ' Determine the set of slides to be split: selected slides (if any)
    ' or the whole slide deck.
    If ActiveWindow.Selection.Type = ppSelectionSlides Then
        split_selected_slides = MsgBox(prompt:="It seems that a set of slides is currently selected. " + _
             "By proceeding, you will only be splitting selected slides." + Chr$(13) + _
             "- Click " + Chr$(34) + "Yes" + Chr$(34) + " if this is what you want." + Chr$(13) + _
             "- Click " + Chr$(34) + "No" + Chr$(34) + " if you want to split ALL the slides in the presentation instead." + Chr$(13) + _
             "- Click " + Chr$(34) + "Cancel" + Chr$(34) + " to simply cancel the operation.", buttons:=vbYesNoCancel, Title:="PPspliT - Information request")
        If split_selected_slides = vbNo Then
            Set slide_range = ActiveWindow.Presentation.Slides.Range
        ElseIf split_selected_slides = vbCancel Then
            Exit Sub
        Else
            Set slide_range = ActiveWindow.Selection.SlideRange
        End If
    Else
        Set slide_range = ActiveWindow.Presentation.Slides.Range
    End If
    
        
    
    ' After a few pre-processing steps, slides are split as follows.
    ' Assume to start from the following slide range ("Anim #" are animation effects in the
    ' timeline of each slide; for simplicity here it is assumed that each effect is mouse-triggered):
    ' +---------+ +---------+ +---------+
    ' | Slide 1 | | Slide 2 | | Slide 3 |
    ' |         | |         | |         |
    ' | Anim #1 | | Anim #1 | | Anim #1 |
    ' | Anim #2 | |         | | Anim #2 |
    ' | Anim #3 | |         | |         |
    ' +---------+ +---------+ +---------+
    '
    ' The following actions are taken for each slide.
    ' First of all, each slide is duplicated as many times as the number of animation effects in
    ' its timeline, plus 1 (representing the initial state of the slide).
    ' After each duplication, emphasis effects are applied, so that their result is persistent in
    ' subsequent copies of each slide:
    ' +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+
    ' | Slide 1 | | Slide 1 | | Slide 1 | | Slide 1 | | Slide 1 | | Slide 2 | | Slide 2 | | Slide 2 | | Slide 3 | | Slide 3 | | Slide 3 | | Slide 3 |
    ' |         | | Copy 0  | | Copy 1  | | Copy 2  | | Copy 3  | |         | | Copy 0  | | Copy 1  | |         | | Copy 0  | | Copy 1  | | Copy 2  |
    ' | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 |
    ' | Anim #2 | | Anim #2 | | Anim #2 | | Anim #2 | | Anim #2 | |         | |         | |         | | Anim #2 | | Anim #2 | | Anim #2 | | Anim #2 |
    ' | Anim #3 | | Anim #3 | | Anim #3 | | Anim #3 | | Anim #3 | |         | |         | |         | |         | |         | |         | |         |
    ' +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+
    '
    ' Notice that the original instance of each slide is still preserved. This is required because
    ' subsequent operations will affect the contents of animation timelines, thus affecting iterators
    ' that will access these timelines.
    ' At this point, each "Copy" corresponds to an animation step in the timeline (which, again, is
    ' preserved unaltered in the original slide). Therefore, for each step the following sets of
    ' shapes are deleted from the corresponding "Copy" slide:
    ' 1) shapes that have an entry effect applied in the future (i.e., they are supposed to appear
    '    at a future step of the timeline)
    ' 2) shapes to which an exit effect was last applied (i.e., they are supposed to have disappeared)
    '
    ' After this processing the original slide, which is no longer needed, is simply removed,
    ' resulting in the final sequence of split slides:
    ' +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+
    ' | Slide 1 | | Slide 1 | | Slide 1 | | Slide 1 | | Slide 2 | | Slide 2 | | Slide 3 | | Slide 3 | | Slide 3 |
    ' | Copy 0  | | Copy 1  | | Copy 2  | | Copy 3  | | Copy 0  | | Copy 1  | | Copy 0  | | Copy 1  | | Copy 2  |
    ' | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 | | Anim #1 |
    ' | Anim #2 | | Anim #2 | | Anim #2 | | Anim #2 | |         | |         | | Anim #2 | | Anim #2 | | Anim #2 |
    ' | Anim #3 | | Anim #3 | | Anim #3 | | Anim #3 | |         | |         | |         | |         | |         |
    ' +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+ +---------+
    '

    ProgressForm.OverallBar.Width = CSng(0)
    ProgressForm.infoLabel = ""
    ProgressForm.Show
    ProgressForm.Repaint
    DoEvents
    
    ' Make sure we are in the correct view mode
    If ActiveWindow.ViewType <> ppViewSlide And ActiveWindow.ViewType <> ppViewNormal Then
        ActiveWindow.ViewType = ppViewNormal
    End If

    ' Bake slide numbers, if asked to
    If slideNumbersAdjustMode <> SLIDENUMBER_DONOTHING Then bakeSlideNumbers slide_range
    
    Dim current_slide As slide, current_original_slide As slide
    Dim effect_sequence As Sequence
    Dim current_effect As effect
    Dim original_slide_count As Integer, effect_count As Integer
    Dim processed_slides_count As Integer, processed_effects_count As Integer
    
    Dim shapeVisibility As Collection
    
    original_slide_count = slide_range.Count
    current_slide_count = original_slide_count
    processed_slides_count = 0
    
    ' Preserve the actual slide number for future usage
    For Each current_slide In slide_range
        current_slide.Tags.Add "originalSlideNumber", Str$(current_slide.SlideIndex)
    Next current_slide


    ' Iterate over all the slides in the presentation
    For Each current_original_slide In slide_range
        current_original_slide.Tags.Delete "done"
        split_slides = 0
    
        processed_slides_count = processed_slides_count + 1
        ProgressForm.SlideNumber = "Slide" + Str$(current_original_slide.Tags("originalSlideNumber")) + " (currently" + Str$(current_original_slide.SlideNumber) + ") -" + Str$(processed_slides_count) + " of" + Str$(original_slide_count)
        DoEvents
        
        preprocessEffects current_original_slide, final_colors
        
        If current_original_slide.timeline.MainSequence.Count > 0 Then
        
            ' There are entry/emphasis/exit effects to process
            
            ProgressForm.infoLabel = "Preprocessing animation effects..."
            DoEvents
        
            ' Preserve shape IDs in case they are lost
            copyShapeIDsToTags current_original_slide
            
            Set effect_sequence = current_original_slide.timeline.MainSequence
            effect_count = effect_sequence.Count
            processed_effects_count = 0
            Set current_slide = current_original_slide
            
            ' shapeVisibility is a dictionary that stores, for each shape (or text paragraph),
            ' its visibility status at the current step of the animation timeline. Here the
            ' data structure is initialized with the IDs of all the shapes/paragraphs involved
            ' in the timeline. At the same time, here we determine the visibility status of
            ' each shape before any animations are played.
            Set shapeVisibility = New Collection
            Dim fullShapeID As String
            Dim final_colors_for_slide As Collection
            For Each current_effect In effect_sequence
                fullShapeID = getFullShapeID(current_effect)
                ' Inserting twice the same key may raise an error
                On Error Resume Next
                shapeVisibility.Add Null, fullShapeID
                On Error GoTo error_handler
                
                ' Determine the initial visibility status
                If IsNull(shapeVisibility(fullShapeID)) Then
                    ' Visibility was undetermined so far, so this is the first effect in the
                    ' animation timeline that is applied to this shape/paragraph
                    shapeVisibility.Remove (fullShapeID)
                    shapeVisibility.Add isEmphasisEffect(current_effect) Or (current_effect.Exit = msoTrue), fullShapeID
                End If
                If cancelStatus Then
                    Unload ProgressForm
                    Exit Sub
                End If
            Next current_effect


            ProgressForm.infoLabel = "Duplicating slides and applying emphasis/path effects..."
            DoEvents
            
            ' Create first duplicated slide ("Copy 0"), which will contain
            ' shapes in their initial state
            Set current_slide = current_slide.Duplicate(1)
            current_slide_count = current_slide_count + 1

            ' Process emphasis effects first, so that they are made persistent across
            ' copies of the original slide
            For Each current_effect In effect_sequence
                If (Not splitMouseTriggered) Or isMouseTriggered(current_effect) Then
                    ' Either the split has been requested for every animation step, or this
                    ' is a click-triggered animation effect.
                    ' Create a copy of the slide and consider the copy as the base for future
                    ' duplications
                    Set current_slide = current_slide.Duplicate(1)
                    current_slide_count = current_slide_count + 1
                End If
                
                If isEmphasisEffect(current_effect) Then
                    Set final_colors_for_slide = final_colors(Str$(current_original_slide.SlideID))
                    applyEmphasisEffect effect_sequence, current_effect, findShapeByIDinTag(current_effect.Shape.Id, current_slide), final_colors_for_slide(Str$(current_effect.Index))
                End If
                
                processed_effects_count = processed_effects_count + 1
                setProgressBar processed_slides_count - 1 + processed_effects_count / (2 * effect_count), original_slide_count
                If cancelStatus Then
                    Unload ProgressForm
                    Exit Sub
                End If
            Next current_effect


            ProgressForm.infoLabel = "Processing entry/exit effects..."
            DoEvents
            
            ' Go again through the animation steps and delete shapes/paragraphs
            ' according to the animation timeline. Note that this
            ' operation cannot be performed earlier, as shapes would
            ' otherwise be lost across slide duplicates and need tos
            ' be restored somehow (for example by copy-paste).
            processed_effects_count = 0
            ' Set current_slide to the first generated duplicate ("Copy 0")
            Set current_slide = current_slide.Parent.Slides(current_original_slide.SlideNumber + 1)
            
            split_slides = split_slides + 1
            augmentSlideNumbers current_slide, split_slides
            
            For Each current_effect In effect_sequence
                If (Not splitMouseTriggered) Or isMouseTriggered(current_effect) Then
                    purgeInvisibleShapes shapeVisibility, effect_sequence, current_slide
                    purgeEffects current_slide
                    ' Mark current slide as completely processed
                    current_slide.Tags.Add "done", "1"
                    If current_slide.SlideNumber < ActivePresentation.Slides.Count Then
                        Set current_slide = current_slide.Parent.Slides(current_slide.SlideNumber + 1)
                        
                        split_slides = split_slides + 1
                        augmentSlideNumbers current_slide, split_slides
                    End If
                End If
                ' Update the actual visibility status of the current shape/paragraph. Note
                ' that emphasis/path motion effects have no impact on shape visibility
                If Not isEmphasisEffect(current_effect) Then
                    shapeVisibility.Remove (getFullShapeID(current_effect))
                    shapeVisibility.Add (current_effect.Exit = msoFalse), getFullShapeID(current_effect)
                End If
                
                processed_effects_count = processed_effects_count + 1
                setProgressBar processed_slides_count - 0.5 + processed_effects_count / (2 * effect_count), original_slide_count
                If cancelStatus Then
                    Unload ProgressForm
                    Exit Sub
                End If
            Next current_effect
            
            If current_slide.Tags("done") <> "1" Then
                ' This slide has to be processed yet
                purgeInvisibleShapes shapeVisibility, effect_sequence, current_slide
                purgeEffects current_slide
            End If
                
            current_original_slide.Delete
        End If
        
    Next current_original_slide
                        
    Unload ProgressForm
    Exit Sub
    
error_handler:
    resp = MsgBox("Unfortunately, an unrecoverable error has occurred while splitting." & vbCrLf & _
                  "- Error code: " & Str$(Err.Number) & vbCrLf & _
                  "- Error description: " & Err.Description & vbCrLf & _
                  "- Slide number: " & current_original_slide.Tags("originalSlideNumber") & " (original) - " & Str$(current_original_slide.SlideNumber) & " (actual)" & vbCrLf & _
                  "Would you like to continue anyway (discouraged)?", vbYesNo, "Fatal error")
    If resp = vbYes Then
        Resume Next
    Else
        On Error GoTo 0
        Resume
    End If
End Sub

' The "Adjust slide numbers" combo box
' has changed its state
Sub ASNdDownChanged(ByRef box As IRibbonControl, ByRef dropDownID As String, ByRef selectedIndex As Variant)
    slideNumbersAdjustMode = selectedIndex
End Sub

' The "Adjust slide numbers" combo box
' is set to "Yes" by default
Sub ASNdDownDefault(ByRef box As IRibbonControl, ByRef selectedItem As Variant)
    selectedItem = SLIDENUMBER_BAKE
    slideNumbersAdjustMode = SLIDENUMBER_BAKE
End Sub

' The "Split on click-triggered animations" check box
' has been clicked
Sub CTcBoxChanged(button As IRibbonControl, pressed As Boolean)
    splitMouseTriggered = pressed
End Sub

' The "Split on click-triggered animations" check box
' is checked by default
Sub CTcBoxDefault(button As IRibbonControl, ByRef state)
    state = True
    splitMouseTriggered = True
End Sub

' Display the about dialog
Sub displayAboutForm()
    AboutForm.Show
End Sub
