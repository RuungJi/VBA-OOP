Attribute VB_Name = "main"
'@Folder("CH1.Duck")
Option Explicit

Sub DuckSimulator()
    Dim Mallard As IDuck
    Set Mallard = New MallardDuck
    Mallard.quack
    Mallard.swim
    Mallard.display
    
    Debug.Print ""

    Dim RedHead As New RedHeadDuck
    RedHead.quack
    RedHead.swim
    RedHead.display
    
    Debug.Print ""
    
    Dim Rubber As IDuck
    Set Rubber = New RubberDuck
    Rubber.quack
    Rubber.swim
    Rubber.display
    
    Debug.Print ""
    
    Dim Decoy As IDuck
    Set Decoy = New DecoyDuck
    Decoy.quack
    Decoy.swim
    Decoy.display
End Sub
