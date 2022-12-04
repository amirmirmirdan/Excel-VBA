Sample structure to get the time for the subroutine to run.

```vb
Sub TimeTestASubroutine()
    Dim StartTime As Date
    Dim EndTime As Date
    
	'   Store the starting time
    StartTime = Timer
	'   Perform some procedures here
	   
	'   Get ending time
    EndTime = Timer
	
	'   Display total time in seconds
    MsgBox Format(EndTime - StartTime, "0.0")
End Sub
```



### Sections
#### Initialize
Timer
```vb

StartTime = Time

```

#### TimeTaken
This function will enable to record the time spent from the start process to a multiple specific part of the process.
Which enables users to make proper evaluations on which code in the process is consuming the most of the time in the process. 
Modification, upgrade, improvement and enhancements opportunities can easily recognise.

```vb

TimeTaken = Timer - StartTime

```

#### Terminate


```vb

EndTime = Timer - StartTime

``