## Excel Formula to use again

## "COLUMNS" is dynamic to allow change referencing multiple columns. https://tinyurl.com/mpwfwdc3 ##
=VLOOKUP($B123,Sheet2!$B$4:$BY$49,COLUMNS($EU$5:GD6)+1,FALSE)

## Convert any measurement (Reference page https://tinyurl.com/yc62xbm4)
=CONVERT(number, “from unit “,”to unit”)

## Remove numbers from String
=IFERROR(SUBSTITUTE(C24,LEFT(C24,MIN(IFERROR(FIND({0,1,2,3,4,5,6,7,8,9},C24),""))-1),""),"")

"%userprofile%\AppData\Local\Virtual Store\Program Files (x86)\JDA\Intactix\ProSpace"
"%userprofile%\AppData\Local\Virtual Store\Program Files (x86)\JDA\Intactix\ProSpace\ProSpace.exe"
