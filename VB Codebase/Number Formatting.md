[[Base|Home]]

---
#### Format Function
```vb
	MsgBox Format(1234567.89, "#,##0.00")
```

This will return = **1,234,567.89**

#### Using Format String
```vb
Range("A1:A7").NumberFormat = "[>=10000][Green]#,##0.00;[<10000][Red]#,##0.00"
```

#### Predefined Formats
##### General Number
```vb
MsgBox Format(1234567.89, "General Number")
```
The result will be 1234567.89

---
##### Currency
```vb
MsgBox Format(1234567.894, "Currency")
```

This format will add a currency symbol in front of the number e.g. $, £ depending on your locale, but it will also format the number to 2 decimal places and will separate the thousands with commas.

The result will be $1,234,567.89

---
##### Fixed
```vb
MsgBox Format(1234567.894, "Fixed")
```

This format displays at least one digit to the left but only two digits to the right of the decimal point.
The result will be 1234567.89

---
##### Standard
```vb
MsgBox Format(1234567.894, "Standard")
```

This displays the number with the thousand separators, but only to two decimal places.
The result will be **1,234,567.89**

---
##### Percent
```vb
MsgBox Format(1234567.894, "Percent")
```

The number is multiplied by 100 and a percentage symbol (%) is added at the end of the number.  The format displays to 2 decimal places
The result will be **123456789.40%**


### Dates & Time 
#### General Date

```vb
MsgBox Format(Now(), "General Date")
```

This will display the date as date and time using AM/PM notation.  How the date is displayed depends on your settings in the Windows Control Panel (Clock and Region | Region). It may be displayed as ‘mm/dd/yyyy’ or ‘dd/mm/yyyy’

The result will be ‘7/7/2020 3:48:25 PM’

#### Long Date

```vb
MsgBox Format(Now(), "Long Date")
```

This will display a long date as defined in the Windows Control Panel (Clock and Region | Region).  Note that it does not include the time.

The result will be ‘Tuesday, July 7, 2020’

#### Medium Date

```vb
MsgBox Format(Now(), "Medium Date")
```

This displays a date as defined in the short date settings as defined by locale in the Windows Control Panel.

The result will be ’07-Jul-20’

#### Short Date

```vb
MsgBox Format(Now(), "Short Date")
```

Displays a short date as defined in the Windows Control Panel (Clock and Region | Region). How the date is displayed depends on your locale. It may be displayed as ‘mm/dd/yyyy’ or ‘dd/mm/yyyy’

The result will be ‘7/7/2020’

#### Long Time

```vb
MsgBox Format(Now(), "Long Time")
```

Displays a long time as defined in Windows Control Panel (Clock and Region | Region).

The result will be ‘4:11:39 PM’

#### Medium Time

```visual-basic
MsgBox Format(Now(), "Medium Time")
```

Displays a medium time as defined by your locale in the Windows Control Panel. This is usually set as 12-hour format using hours, minutes, and seconds and the AM/PM format.

The result will be ’04:15 PM’

#### Short Time

```visual-basic
MsgBox Format(Now(), "Short Time")
```

Displays a medium time as defined in Windows Control Panel (Clock and Region | Region). This is usually set as 24-hour format with hours and minutes

The result will be ’16:18’