Attribute VB_Name = "Module1"
Private Type PathType
  Node() As D3DVECTOR
  NumNode As Integer
  CurrentNode As Integer
  MeshAssocied As Integer
  FollowLandscape As Boolean
  Position As D3DVECTOR
  Rotation As D3DVECTOR
  Direction As D3DVECTOR
End Type

Private PathT As PathType


Public Sub AddPathNode(v As D3DVECTOR)
  '##BD Add a path node vector to the node list.
  '##PD V Node to add.
  PathT.NumNode = PathT.NumNode + 1
  ReDim Preserve PathT.Node(PathT.NumNode)
  PathT.Node(PathT.NumNode) = v
End Sub


Function GetNode(NodeIndex As Integer) As D3DVECTOR
   '##BD Returns the position of the specified node.
   '##PD NodeIndex Node Index (range :1-NodeCount)
   GetNode = PathT.Node(NodeIndex)
End Function

Function GetNodeCount() As Integer
  '##BD Returns the nodes count in the current path.
  GetNodeCount = PathT.NumNode
End Function

Private Function GetNodeId(ByVal t As Integer) As D3DVECTOR
   If t < 1 Then
    t = t + PathT.NumNode
   End If
   If t > PathT.NumNode Then
    t = t - PathT.NumNode
   End If
   GetNodeId = PathT.Node(t - 1)
   
End Function

Public Function GetSplinePoint(step As Single, strenth As Single) As D3DVECTOR
  Dim NextStep As Integer
  Dim CurrentStep As Integer, S As Single
  Dim t2 As Single
  Dim t3 As Single
  Dim M(4) As Single
  Dim temp As D3DVECTOR
  Dim i As Integer
  Dim t As Single
'  Dim T1 As D3DVECTOR
'  Dim t2 As D3DVECTOR
  Dim Out As D3DVECTOR
  CurrentStep = Int(step)
  t = step - CurrentStep
  
 ' T1 = (VSubtract(GetNodeId(CurrentStep), GetNodeId(CurrentStep - 2)))
 ' t2 = (VSubtract(GetNodeId(CurrentStep + 3), GetNodeId(CurrentStep + 1)))
 ' D3DXVec3Hermite Out, GetNodeId(CurrentStep), T1, GetNodeId(CurrentStep + 1), t2, S

'  GetSplinePoint = Out

   Dim Ret As D3DVECTOR
    Ret.x = 0
    Ret.y = 0
    Ret.z = 0
    
    
    t2 = t * t '* strenth
    t3 = t * t * t
    
    
    M(0) = (0.5 * ((-1# * t3) + (2# * t2) + (-1# * t)))
    M(1) = (0.5 * ((3# * t3) + (-5# * t2) + (0# * t) + 2#))
    M(2) = (0.5 * ((-3# * t3) + (4# * t2) + (1# * t)))
    M(3) = (0.5 * ((1# * t3) + (-1# * t2) + (0# * t)))

  '  Dim ret As D3DVECTOR
    For i = 0 To 3
        'MsgBox (temp & " " & GetNodeId(CurrentStep + i) & " " & M(i))
        D3DXVec3Scale temp, GetNodeId(CurrentStep + i), M(i)
        D3DXVec3Add Ret, Ret, temp
        
    Next i
    
  GetSplinePoint = Ret

End Function

Public Sub ResetPath()
  '##BD Reset the path settings and clears all the nodes.
  PathT.NumNode = 0
  ReDim PathT.Node(0)
End Sub
