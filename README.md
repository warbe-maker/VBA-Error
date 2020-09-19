# Common VBA Error Handler
### Coverage
The error handling approach significantly differs between development/test and production.
#### Development and Test
- Debug.Print of the Error Description
- Stop when the error occurs within the procedure it occurred
- Manual _Resume_ of the code line which caused the error providing the chance to change the code on the fly - what somebody called a "godsend" when needed.

Because this for sure is something unwanted in production it will be turned off in production by the Conditional Compile Argument _Debugging=0_
#### Production
When the [_Entry Procedure_](#the-entry-procedure) is known, the error is passed on back up to it and finally displays a message with:
- A title which considers whether it is an _Application Error_ or a _Visual Basic Run Time Error_ in the form<br>Application|VB error [number] in <module>.<procedure> [at line <line number>] 
- The _Error Description_ which is either the description of the _Visual Basic Run Time Error_ (```Err.Description```) or the description of the error provided by means of the ```Err.Raise``` statement 
- The _Error Path_ (provided the [_Entry Procedure_](#the-entry-procedure) is known) indicating the call stack from the procedure where the error occurred (the _Error Source_ back up to the [_Entry Procedure_](#the-entry-procedure)


Provided the  Conditional Compile Argument _ExecTrace=1_, each time an _Entry Procedure_ is reached, the _Execution Trace_  including the _Execution Time_ of each [traced procedure](#execution-traced-procedures) is printed in the VBE immediate window.

Note: When the [_Entry Procedure_](#the-entry-procedure) is unknown  the error is immediately displayed in the procedure where the error occurred or in the first calling procedure which has an ```On Error Goto ...``` statement.

### Basic error handling
The below approach is absolutely recommendable. Its wise to have it coded before an error occurs but can be added then as well. :

```vbscript
Private Sub Any
   On Error Go-to on_error
   ....
   
exit_proc:
   Exit Sub
   
on_error:
#If Debugging Then
   Debug.Print Err.Description: Stop: Resume
#End If
End Sub
```
Of course this is absolutely inappropriate when the project runs productive. The above should only be active for development and test and the complete common error handling should run in production as follows:
```vbscript
Private Sub Any
   On Error Go-to on_error
   ....
   
exit_proc:
   Exit Sub
   
on_error:
#If Test Then
   Debug.Print Err.Description: Stop: Resume
#Else
   <the production error banking see chapter Usage>
#End If
End Sub
```
### Installation
Download and import to you VBA project:
- mErrHndlr
- clsExecTrace
- clsCallStack
### Usage
1. The identification of the _Entry Procedure_ is the key to some of the key features of the _Common VBA Error Handler_. It requires  a BoP Begin of Procedure and a EoP (End of Procedure statement.
2. The identification of the _Error Source_ requires a ```Const PROC ="...." ``` statement in each procedure which has an ```On Error Goto on_error``` statement and the following function copied into the module:
```vbscript
Private Property ErrSrc(Optional ByVal s As String) As String
    ErrSrc = "mTest" & "." & s
End Function
```
The usage of the _Common VBA Error Handler_  in a procedure with an error handling will look as follows:

```vbscript
Private Sub Any
   On Error Goto on_error
   Const PROC = "Any" ' identifies this procedure as error source in case
   BoP ErrSrcPROC) ' Begin of procedure, mandatory in an "entry procedure"
   
   <code>
   
exit_proc:
   EoP ErrSrc(PROC) ' End of procedure, mandatory in an "entry procedure"
   Exit Sub
   
on_error:
#If Test Then
   Debug.Print Err.Description: Stop: Resume
#Else
   mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
#End If
End Sub
```
#### Execution traced procedures