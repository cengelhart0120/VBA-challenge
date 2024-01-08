# VBA-challenge
Screenshots, VBA script files, and a README file submitted for the completion of Module 2 Challenge

For the most part, the written VBA code was derived from curriculum content, via instruction and/or in-class activities, but there were two key sources consulted for tying everything together.

## File "VBA_challenge.bas"

Contents of this macro were derived from curriculum content, via instruction and/or in-class activities.

## File "VBA_challenge_loop.bas"

Contents of this macro were derived from two sources, because two things were needed:
    1. A way to run "VBA_challenge.bas" across all worksheets in the workbook; and
    2. A way to run the "VBA_challenge.bas" macro from within the For loop established in #1 above.
    
The first source was found at:
    http://www.vbaexpress.com/forum/showthread.php?64458-How-to-run-macro-for-multiple-worksheets
Specifically from the entry on 01-18-2019 at 02:10 PM. I didn't initially understand the purpose of specific syntax choices, but figured the macro above was somehow referring to the macro below.

This led to seeking and finding the second source at:
    https://www.excelcampus.com/vba/vba-call-statement-run-macro-from-macro
This page more explicitly outlined the use of a "Call" statement for running a macro within a different macro (and cleared up confusion about the lack of a "Call" statement in the first source, which isn't needed, but this author encouraged its use).

## File "VBA_challenge_combined.bas"

This is the combination of the above two macros in a single .bas file that was used to complete the challenge.
