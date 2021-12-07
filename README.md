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
    On Error Goto eh
    ' .....
eh: mErH.ErrMsg error-source[, buttons]
```
The _ErrMsg_ service has these named arguments:

|  Argument   |   Description   |
| ----------- | --------------- |
| err_source  | Obligatory, string expression providing \<module>.\<procedure> |
| err_buttons | Optional. Variant. Defaults to "Terminate execution" button when omitted.<br>May be a VBA MsgBox like value or any descriptive button caption string (including line breaks for a multi-line caption). The buttons may be provided as a comma delimited string, a collection or a dictionary. vbLf items display the following buttons in a new row (up to 7 rows are available). |

### The _AppErr_ service
In order to not confuse errors raised with `err.Raise ...` the service adds the [_vbObjectError_][7] constant to a given positive number to turn it into a negative. An advantage by the way: Each procedure can have it's own positive error numbers ranging from 1 to n with `err.Raise mErH.AppErr(n)`. The _ErrMsg_ service, when detecting a negative error number uses the _AppErr_ service to turn it back into it's original positive error number.

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

### The _BoTP_ (begin of Test Procedure) service for automating regression tests
An - preferably automated - regression test will execute a series of test procedures. Any interruption other than one caused by a failed assertion n assertion should thus be avoided. The _BoTP_ allows the specification of **asserted error numbers** for procedures testing error conditions. For an error number indicated 'asserted' the _ErrMsg_ service bypassed the display of the error message.

The BoTP service has the following Syntax:<br>
`BoTP procedure-id, err-number[, err-number] ...`
with the following named arguments:

|      Argument     |   Description                                                   |
| ----------------- | --------------------------------------------------------------- |
| botp_id           | Obligatory, Expression providing a unique name of the procedure |
| botp_errs_asserted| Obligatory, ParamArray with positive numbers                    |

## Installation
- Download and import the module  [_mErH_][1]
- Download the UserForm [fMsg.frm][2] and [fMsg.frx][3] and import _fMsg.frm_
- Download and import [mMsg.bas][4]
- Since the extra effort is very little, by the way installing the _Common VBA Execution Trace Service_ is worth being considered:<br>Download [mTrc.bas][5] and import it.

## Usage
See blog-post [A common VBA Error Handler][6] for the details

## Contribution, development, test, maintenance
The dedicated _Common VBA Component Workbook_ **[ErH.xlsm][8]** is used for development, test, and maintenance. I keep this Workbook in a dedicated folder which is the local equivalent (in github terminology the clone of this public GitHub repo. The module **_mTest_** contains all obligatory test procedures executed when the code is modified. Code modifications are preferrably made in a Github branch which is merged to the master once a code change has finished and successfully tested.

Those interested not only in using the _Common VBA Error Services_ but also feel prepared to ask question, make suggestions, open raising issues may fork or clone the [public repo][8] to their own computer. A process which is very well supported by the [GitHub Desktop for Windows][9] which is the environment I do uses for the version control and a a continuous improvement process.

[1]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/mErH.bas
[2]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/source/fMsg.frm
[3]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/source/fMsg.frx
[4]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/source/mMsg.bas
[5]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Message-Service/master/source/mTrc.bas
[6]:https://warbe-maker.github.io/warbe-maker.github.io/vba/common/error/handling/2021/01/16/Common-VBA-Error-Services.html
[7]:https://docs.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.constants.vbobjecterror?view=netcore-3.1
[8]:https://gitcdn.link/repo/warbe-maker/Common-VBA-Error-Services/master/source/ErH.xlsm
