# Common VBA Error Handler
### Coverage
What the error handling provides differs between development and test and production
#### Development and Test
- Debug.Print of the Error Description
- Stop when the error occurs within the procedure it occurred
- Manual _Resume_ of the code line which caused the error providing the chance to change the code on the fly.
- 
#### Production
When the _Entry Procedure is known, the error is passed on back up to it and finally the error is displayed with:
- The error number with the distinction of an _Application Error_ from a _Visual Basic Run Time Error_
- The _Error Description_ which is either the description of the _Visual Basic Run Time Error_ or the description of the error provided by means of the Err.Raise statement 
- The _Error Source_ in the form  <module>.<procedure>
- The _Error Line_ provided the procedure where the error occurred has line numbers
- The _Error Path_ as the call stack from the procedure where the error occurred back up to the _Entry Procedure_ provided it is known
- Optionally, each time when an _Entry Procedure_ is reached an _Execution Trace which includes the _Execution Tine_ of the traced procedures.

Note: When the [_Entry Procedure_](#the-entry-procedure) is unknown  the error is immediately displayed in the procedure where the error occurred or in the calling procedure which has an On Error Goto statement

### Basic error handling
When an error occurred during development/test the best what can happen is a stop with the chance to re-execute the line which caused the error and that happens with:

```vbscript
Private Sub Any
   On Error Go-to on_error
   ....
   
exit_proc:
   Exit Sub
   
on_error:
   Debug.Print Err.Description: Stop: Resume
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
 - The identification of the _Entry Procedure_ is the key to some of the key features of the _Common VBA Error Handler_. It requires  a BoP Begin of Procedure and a EoP (End of Procedure statement.
 - For the identification of the _Error Source_ have a ```Const PROC ="...." ``` statement in each procedure which has an ```On Error Goto on_error``` statement and copy the following function into each module:
```vbscript
```
The usage of the _Common VBS Error Handler_  in a procedure with an error handling will look as follows:

```vbscript
Private Sub Any
   On Error Go-to on_error
   Const PROC = "Any" ' identifies this procedure as error source in case
   BoP ErrSrcPROC) ' Begin of procedure
   ....
   
exit_proc:
   EoP ErrSrc(PROC) ' End of procedure
   Exit Sub
   
on_error:
#If Test Then
   Debug.Print Err.Description: Stop: Resume
#Else
   mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
#End If
End Sub
```