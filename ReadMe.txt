Pocket Calculator using a State Transition Table

This pocket calculator is meant as an example for the almost forgotten State Transition Table technique. With this technique the complete behaviour pattern is stored in a two-dimensional table containig Reactions to input-triggers (in this case key presses, but depending on circumstances the triggers may be anything else), and States. An STT is an ideal tool to control game characters for example because it allows different reactions to identical inputs.

The code is documented and also contains a short description of how it works. It's worth the download if you want to learn more about STTs, so give it a try. Download is 18 kB.

Rename Gradient.ocx.dat as Gradient.ocx and register it using RegSvr32 or else remove all references to graRainbow.


How to use the calculator examples:
-----------------------------------

Normal calculations            4 * 7 = 
Chain calculations (a)         4 * 5 / 7 - 3 =
Chain calculations (b)         8 - 4 = + 3 =
Omitted operand                12 * =
Roots (eg cube root)           3 R 8 =    
Exponentiation                 7 E 5 =
Memory operations              M + | M -  etc. 
Recall memory                  M =
Clear                          C
Clear operand                  C
Clear operator                 C
Clear memory                   M C
Exchange operands              X
Exchange operand with memory   M X
Toggle sign                    P
Toggle sign of memory          M P    
Copy result to clipboard       Ctrl with C

Case (upper/lower) of letters is irrelevant.
You can also use the numeric key pad.                            
Error texts are taken from the Err object.