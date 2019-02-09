VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'peek, pop, push, IsEmpty

private topNode as Node

function Pop() as Object
  Pop = topNode.Value
  set topNode = topNode.PreviousNode
  set topNode.NextNode = nothing
end function

function Peek() as Object
  set Peek = topNode.Value
end function

function Push(obj as Object)
  dim newNode as Node: set newNode = new Node
  set newNode.Value = obj
  set topNode.NextNode = newNode
  set topNode = newNode
end function

function IsEmpty() as bool
  if topNode is nothing then IsEmpty = True
end function