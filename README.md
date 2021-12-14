# Common VBA Error Services

## Services
### The _ErrMsg_ service
displays a well structured error message with
  - the **type of the error** (Application error, VB Runtime error, or Database error) 
  - the **description** of the error (_err.Description_),
  - the **error source**,
  - the **path to the error** (provided the **_Entry Procedure_** is known, 
  - an optional **additional information about an error**
  - **free specified buttons**
  - the **error line** when available
- waits for the user's button clicked and provides/returns the reply button's value for processing by the caller<br>(obsolete of course when on only the default button is displayed).

The _ErrMsg_ service has the following syntax (error description and error line are obtained from the err object)
```VB
    Const PROC = "<procedure-name>"
    
    On Error Goto eh
    ' .....
eh: ErrMsg ErrSource(PROC)
```

With the debugging option, activated by the Conditional Compile Argument 'Debugging = 1' debugging becomes as quick and easy as possible:

```VB
    Const PROC = "<procedure-name>"
    
    On Error Goto eh
    ' .....
xt: Exit Sub/Function 

eh: Select Case ErrMsg(ErrSource(PROC))
      Case vbResume:  Stop: Resume:
      Case vbResumeNext:  Stop: Resume Next
      Case Else:          Goto xt ' clean exit
    End Select
```

### The _AppErr_ service
In order to not confuse errors raised with `Err.Raise ...` the service adds the [_vbObjectError_][7] constant to a given positive number to turn it into a negative number. I.e. each procedure can have it's own positive error numbers ranging from 1 to n. The _ErrMsg_ service considers a negative error number an Application Error (in contrast to a Runtime Error) and uses the _AppErr_ service to turn it back into its original positive error number.

### The _BoP/EoP_ path to the error service
The _ErrMsg_ service only displays a path to the error when an _Entry Procedure_ had been indicated. The path to the error is assembled when the error is passed on from the error source up to the _Entry Procedure_ where the error is displayed when reached.

The _BoP/EoP_ services have the following syntax:<br>
`mErH.BoP procedure-id[, arguments]`<br>
`mErH.EoP procedure-id`<br>
with the following named arguments:

| Service |   Argument   |   Description                                                                             |
| ------- | ------------ | ----------------------------------------------------------------------------------------- |
| BoP     | bop_id       | Obligatory, String expression, unique identification of the procedure in the module       |
| BoP     | bop_arguments| Optional, ParamArray, a list of the procedures argument, optionally paired as name, value |
| EoP     | eop_id       | Obligatory, String expression, unique identification of the procedure name in the module  |

Note: When the user not only has one reply button but several reply choices (see the debugging service for instance), the error message is displayed immediately with the procedure which caused the error. In this case the path to the error is composed from a stack which is maintained along with each BoP/EoP statement. I.e. the path to the error contains only procedures which do use BoP/EoP statements.

### The debugging service for identifying an error line
With the _Conditional Compile Argument_ `Debuggig = 1` the error message is displayed with two additional buttons which allow a `Stop: Resume` reaction which leads to the code line the error occurred.

### The _BoTP_ service
An - preferably automated - regression test will execute a series of test procedures. Any interruption other than one caused by a failed assertion should thus be avoided. The _BoTP_ allows the specification of **asserted error numbers** for procedures which do test error conditions. For any asserted error number the _ErrMsg_ display service is bypassed.

The BoTP service has the following Syntax:<br>
`BoTP procedure-id, err-number[, err-number] ...`
with the following named arguments:

|      Argument     |   Description                                                   |
| ----------------- | --------------------------------------------------------------- |
| botp_id           | Obligatory, Expression providing a unique name of the procedure |
| botp_errs_asserted| Obligatory, ParamArray with positive numbers                    |

### The _Regression_ service
With `mErH.Regression = True` the _ErrMsg_ service runs in regression test mode - and thus is used only in a regression test procedure. As a result any error specifically tested will not be displayed in order to keep the regression test uninterrupted. Which error numbers are regarded asserted (tested, anticipated respectively) is specified with the _[BoTP](#the-botp-service)_ service.

## Installation
- Download and import the standard module  [_mErH.bas_][1]
- Download the UserForm [fMsg.frm][2] and the standard module [fMsg.frx][3] and import _fMsg.frm_
- Download and import the standard module [mMsg.bas][4]
- Optionally download the standard module [mTrc.bas][5] and import it. With very little extra effort the _Common VBA Execution Trace Service_ becomes available for the VBA-Project by the way.

## Usage
See blog-post [A common VBA Error Handler][6] for the details

## Contribution, development, test, maintenance
The dedicated _Common VBA Component Workbook_ **[ErH.xlsm][8]** is used for development, test, and maintenance. I keep this Workbook in a dedicated folder which is the local equivalent (in github terminology the clone of this public GitHub repo. The module **_mTest_** contains all obligatory test procedures executed when the code is modified. Code modifications are preferably made in a Github branch which is merged to the master once a code change has finished and successfully tested.

Those interested not only in using the _Common VBA Error Services_ but also feel prepared to ask question, make suggestions, open raising issues may fork or clone the [public repo][8] to their own computer. A process which is very well supported by the [GitHub Desktop for Windows][9] which is the environment I do uses for the version control and a a continuous improvement process.

[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas
[2]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/fMsg.frm
[3]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/fMsg.frx
[4]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mMsg.bas
[5]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mTrc.bas
[6]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/error/handling/2021/01/16/Common-VBA-Error-Services.html
[7]:https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1
[8]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/ErH.xlsm
