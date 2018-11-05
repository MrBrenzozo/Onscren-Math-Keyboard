# Onscren-Math-Keyboard

Purpose;
This program sends ASCII keystrokes to a word document. if you have the macro enalbed word docement i made, 
when you press the keyboard button, it will trigger a macro on the word document to paste the math operator into
the word document. 

Room for improvement;
i am trying to figure out how to use the code from the macros in the keyboard program to send VBA oMath objects
directly to the word document so that a macro is not even needed. The macros needed are pasted below, just paste
into module1 and start using it.

'\\Code for Module1
Option Explicit

Sub AddKeyBinding()
    With Application
        '\\Do customization in THIS document
        .CustomizationContext = ThisDocument
        
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKey5)).Disable
        '\\Add keyinding to this document shortcut: Alt+5
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKey5), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="repeating_digit"
    
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKey6)).Disable
        '\\Add keyinding to this document shortcut: Alt+6
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKey6), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="radical"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKey7)).Disable
        '\\Add keyinding to this document shortcut: Alt+7
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKey7), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="exponent"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKey8)).Disable
        '\\Add keyinding to this document shortcut: Alt+8
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKey8), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="brackets"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKey9)).Disable
        '\\Add keyinding to this document shortcut: Alt+9
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKey9), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="reciprocal"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKey0)).Disable
        '\\Add keyinding to this document shortcut: Alt+0
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKey0), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="fraction"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyE)).Disable
        '\\Add keyinding to this document shortcut: Alt+n+e
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyE), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="not_equal"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyT)).Disable
        '\\Add keyinding to this document shortcut: Alt+a+p
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyT), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="approximate"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyU)).Disable
         '\\Add keyinding to this document shortcut: Alt+p+m
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyU), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="plus_or_minus"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyI)).Disable
         '\\Add keyinding to this document shortcut: Alt+p+m
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyI), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="absolute_value"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyO)).Disable
        '\\Add keyinding to this document shortcut: Alt+g+t
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyO), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="greater_than"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyA)).Disable
        '\\Add keyinding to this document shortcut: Alt+l+t
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyA), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="less_than"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyD)).Disable
        '\\Add keyinding to this document shortcut: Alt+g+0
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyD), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="greater_or_equal"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyK)).Disable
        '\\Add keyinding to this document shortcut: Alt+l+0
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyK), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="less_or_equal"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyX)).Disable
        '\\Add keyinding to this document shortcut: Alt+m+b
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyX), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="square_bracket"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyC)).Disable
        '\\Add keyinding to this document shortcut: Alt+k+b
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyC), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="curly_bracket"
               
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyV)).Disable
        '\\Add keyinding to this document shortcut: Alt+g+l
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyV), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="logarithm"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyAlt, wdKeyB)).Disable
        '\\Add keyinding to this document shortcut: Alt+l+m
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyB), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="subscript"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey4)).Disable
        '\\Add keyinding to this document shortcut: Alt+d+g
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey4), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="degree"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey5)).Disable
        '\\Add keyinding to this document shortcut: Alt+k+p
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey5), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="mu"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey6)).Disable
        '\\Add keyinding to this document shortcut: Alt+p+i
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey6), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="pi"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey7)).Disable
        '\\Add keyinding to this document shortcut: Alt+i+f
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey7), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="infinity"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey8)).Disable
        '\\Add keyinding to this document shortcut: Alt+a+g
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey8), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="angle"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey9)).Disable
        '\\Add keyinding to this document shortcut: Alt+t+f
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey9), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="arrow_right"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey0)).Disable
        '\\Add keyinding to this document shortcut: Alt+s+i
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey0), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="sine"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyQ)).Disable
        '\\Add keyinding to this document shortcut: Alt+k+n
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyQ), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="cosine"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW)).Disable
        '\\Add keyinding to this document shortcut: Alt+t+n
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="tangent"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyA)).Disable
        '\\Add keyinding to this document shortcut: Alt+k+t
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyA), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="cotangent"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyG)).Disable
        '\\Add keyinding to this document shortcut: Alt+k+s
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyG), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="cosecant"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyH)).Disable
        '\\Add keyinding to this document shortcut: Alt+s+t
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyH), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="secant"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyJ)).Disable
        '\\Add keyinding to this document shortcut: Alt+t+h
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyJ), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="theta"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyX)).Disable
        '\\Add keyinding to this document shortcut: Alt+r+e
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyX), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="omega"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyB)).Disable
        '\\Add keyinding to this document shortcut: Alt+k+f
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyB), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="ceiling_function"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKey1)).Disable
        '\\Add keyinding to this document shortcut: Alt+f+l
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKey1), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="floor_function"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKey3)).Disable
        '\\Add keyinding to this document shortcut: Alt+k+d
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKey3), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="quadratic_equation"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKey4)).Disable
        '\\Add keyinding to this document shortcut: Alt+p+t
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKey4), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="pythagorean_thereom"
        
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKey6)).Disable
        '\\Add keyinding to this document shortcut: Alt+d+s
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKey6), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="binomial_rule"
        
         '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKey7)).Disable
        '\\Add keyinding to this document shortcut: Alt+f+s
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKey7), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="binomial_rule_factored"
              
        '\\delete existing binding
        FindKey(BuildKeyCode(wdKeyControl, wdKey8)).Disable
        '\\Add keyinding to this document shortcut: Alt+f+s
        .KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKey8), _
        KeyCategory:=wdKeyCategoryMacro, _
        Command:="function_of"
        
        
        
    End With
End Sub
Sub repeating_digit()
'
' repeating_digit Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionAcc).Acc _
        .Char = 773
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub radical()
'
' radical Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionRad).Rad _
        .HideDeg = False
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
End Sub
Sub exponent()
'
' exponent Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSup
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
End Sub
Sub brackets()
'
' brackets Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub reciprocal()
'
' reciprocal Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionFrac). _
        Frac.Type = wdOMathFracBar
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="1"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
End Sub
Sub fraction()
'
' fraction Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionFrac). _
        Frac.Type = wdOMathFracBar
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
End Sub
Sub not_equal()
'
' not_equal Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=8800, Unicode:=True, Bias:=0
End Sub
Sub approximate()
'
' approximate Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=8776, Unicode:=True, Bias:=0
End Sub
Sub plus_or_minus()
'
' plus_or_minus Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=177, Unicode:=True, Bias:=0
End Sub
Sub absolute_value()
'
' absolute_value Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 124
        .Delim.SepChar = 0
        .Delim.EndChar = 124
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub greater_than()
'
' greater_than Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=62, Unicode:=True, Bias:=0
End Sub
Sub less_than()
'
' less_than Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=60, Unicode:=True, Bias:=0
End Sub
Sub greater_or_equal()
'
' greater_or_equal Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=8805, Unicode:=True, Bias:=0
End Sub
Sub less_or_equal()
'
' less_or_equal Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=8804, Unicode:=True, Bias:=0
End Sub

Sub square_bracket()
'
' square_bracket Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 91
        .Delim.SepChar = 0
        .Delim.EndChar = 93
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub curly_bracket()
'
' curly_bracket Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 123
        .Delim.SepChar = 0
        .Delim.EndChar = 125
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub logarithm()
'
' logarithm Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionFunc
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSub
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.TypeText Text:="log"
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    Selection.TypeText Text:="b"
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="x"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="=y"
    Selection.MoveLeft Unit:=wdCharacter, Count:=7
End Sub
Sub omega()
'
' infinity Macro
'
'
    Selection.InsertSymbol Font:="Cambria Math", CharacterNumber:=969, _
       Unicode:=True
End Sub
Sub degree()
'
' degree Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=176, Unicode:=True, Bias:=0
End Sub
Sub coordinate_pair()
'
' coordinate_pair Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:=","
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub pi()
'
' pi Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol Font:="Cambria Math", CharacterNumber:=-10187, _
        Unicode:=True
    Selection.InsertSymbol Font:="Cambria Math", CharacterNumber:=-8437, _
        Unicode:=True
End Sub
Sub infinity()
'
' infinity Macro
'
'
    Selection.InsertSymbol Font:="Cambria Math", CharacterNumber:=8734, _
        Unicode:=True
End Sub
Sub angle()
'
' angle Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol Font:="Cambria Math", CharacterNumber:=8736, _
        Unicode:=True
End Sub
Sub arrow_right()
'
' arrow_right Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=8594, Unicode:=True, Bias:=0
End Sub
Sub sine()
'
' sine Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionFunc
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.TypeText Text:="sin"
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub cosine()
'
' cosine Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionFunc
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.TypeText Text:="cos"
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub tangent()
'
' tangent Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionFunc
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.TypeText Text:="tan"
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub cotangent()
'
' cotangent Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionFunc
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.TypeText Text:="cot"
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub cosecant()
'
' cosecant Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionFunc
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.TypeText Text:="csc"
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub secant()
'
' secant Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionFunc
    Selection.MoveLeft Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.TypeText Text:="sec"
    Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdMove
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub theta()
'
' theta Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.InsertSymbol CharacterNumber:=-10187, Unicode:=True, Bias:=0
    Selection.InsertSymbol CharacterNumber:=-8445, Unicode:=True, Bias:=0
End Sub
Sub reciprocal_expnent()
'
' reciprocal_expnent Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSup
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="-1"
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
End Sub
Sub ceiling_function()
'
' ceiling_function Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 8968
        .Delim.SepChar = 0
        .Delim.EndChar = 8969
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub floor_function()
'
' floor_function Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 8970
        .Delim.SepChar = 0
        .Delim.EndChar = 8971
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
End Sub
Sub quadratic_equation()
'
' quadratic_equation Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.TypeText Text:="x="
    Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionFrac). _
        Frac.Type = wdOMathFracBar
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="-"
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="b"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.InsertSymbol CharacterNumber:=177, Unicode:=True, Bias:=0
    Selection.OMaths(1).Functions.Add(Selection.Range, wdOMathFunctionRad).Rad _
        .HideDeg = True
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSup
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="b"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="2"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="-4ac"
    Selection.MoveDown Unit:=wdLine, Count:=1
    Selection.TypeText Text:="2a"
End Sub
Sub pythagorean_thereom()
'
' pythagorean_thereom Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSup
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="c"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="2"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="="
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSup
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="a"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="2"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="+"
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSup
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="b"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="2"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
End Sub
Sub binomial_rule()
'
' binomial_rule_1_factored Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    Selection.TypeText Text:="-"
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="+"
    Selection.MoveLeft Unit:=wdCharacter, Count:=4
End Sub
Sub binomial_rule_factored()
'
' binomial_rule_factored Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSup
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="2"
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="+2"
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.TypeText Text:="+"
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSup
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.TypeText Text:="2"
    Selection.MoveLeft Unit:=wdCharacter, Count:=15
    Selection.MoveRight Unit:=wdCharacter, Count:=2
End Sub
Sub function_of()
'
' function_of Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.TypeText Text:="f"
    With Selection.OMaths(1).Functions.Add(Selection.Range, _
        wdOMathFunctionDelim, 1)
        .Delim.BegChar = 40
        .Delim.SepChar = 0
        .Delim.EndChar = 41
        .Delim.Grow = True
        .Delim.Shape = wdOMathShapeCentered
    End With
    Selection.TypeText Text:="="
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:="x"
End Sub
Sub subscript()
'
' subscript Macro
'
'
    Selection.OMaths.Add Range:=Selection.Range
    Selection.OMaths(1).Functions.Add Range:=Selection.Range, Type:= _
        wdOMathFunctionScrSub
    Selection.MoveLeft Unit:=wdCharacter, Count:=2
End Sub
Sub mu()
'
' mu Macro
'
'
    Selection.InsertSymbol CharacterNumber:=956, Unicode:=True, Bias:=0
End Sub



