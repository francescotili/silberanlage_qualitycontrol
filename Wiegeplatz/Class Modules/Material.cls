' --------------------------------------------
' MATERIAL CLASS
' Class for the Material data
' --------------------------------------------

Public number As String
Public name As String
Public baseMaterial As String
Public weight As Long
Public plant As String

Public copper As Boolean
Public copperLayer As Long
Public copperBedarf As Long
Public copperEZ As Long
Public copperCurrent As Long

Public silver As Boolean
Public silverLayer As Long
Public silverBedarf As Long
Public silverEZStr As Long
Public silverCurrentStr As Long
Public silverEZGlo As Long
Public silverCurrentGlo As Long

Public xRayMethode As String
Public alterung As Boolean

Private Sub Init( _
  number As String, _
  name As String, _
  baseMaterial As String, _
  weight As Long, _
  plant As String, _
  copper As Boolean, _
  copperLayer As Long, _
  copperBedarf As Long, _
  copperEZ As Long, _
  copperCurrent As Long, _
  silver As Boolean, _
  silverLayer As Long, _
  silverBedarf As Long, _
  silverEZStr As Long, _
  silverCurrentStr As Long, _
  silverEZGlo As Long, _
  silverCurrentGlo As Long, _
  xRayMethode As String, _
  alterung As Boolean _
)
  
  Me.number = number
  Me.name = name
  Me.baseMaterial = baseMaterial
  If weight = "" Then
    Me.weight = 0
  Else
    Me.weight = CLng(weight)
  End If
  Me.plant = plant
  Me.copper = copper
  If copperLayer = "" Then
    Me.copperLayer = 0
  Else
    Me.copperLayer = CLng(copperLayer)
  End If
  If copperBedarf = "" Then
    Me.copperBedarf = 0
  Else
    Me.copperBedarf = CLng(copperLayer)
  End If
  If copperEZ = "" Then
    Me.copperEZ = 0
  Else
    Me.copperEZ = CLng(copperLayer)
  End If
  If copperCurrent = "" Then
    Me.copperCurrent = 0
  Else
    Me.copperCurrent = CLng(copperLayer)
  End If
  Me.silver = silver
  If silverLayer = "" Then
    Me.silverLayer = 0
  Else
    Me.silverLayer = CLng(copperLayer)
  End If
  If silverBedarf = "" Then
    Me.silverBedarf = 0
  Else
    Me.silverBedarf = CLng(copperLayer)
  End If
  If silverEZStr = "" Then
    Me.silverEZStr = 0
  Else
    Me.silverEZStr = CLng(copperLayer)
  End If
  If silverCurrentStr = "" Then
    Me.silverCurrentStr = 0
  Else
    Me.silverCurrentStr = CLng(copperLayer)
  End If
  If silverEZGlo = "" Then
    Me.silverEZGlo = 0
  Else
    Me.silverEZGlo = CLng(copperLayer)
  End If
  If silverCurrentGlo = "" Then
    Me.silverCurrentGlo = 0
  Else
    Me.silverCurrentGlo = CLng(copperLayer)
  End If
  Me.xRayMethode = xRayMethode
  Me.alterung = alterung
End Sub

