Rem	This creates a file 64b in size

echo "This is just a sample line appended to create a big file. " > dummy.txt

Rem This line appends to the file nn number of times where nn is the 3rd number in the Brackets.  Basically it doubles the file size everytime it runs

for /L %%i in (1,1,23) do type dummy.txt >> dummy.txt

Rem	This copies file nn times where nn is the 3rd number in brackets

for /L %%i in (1,1,110) do copy dummy.txt dummy%%i.txt
