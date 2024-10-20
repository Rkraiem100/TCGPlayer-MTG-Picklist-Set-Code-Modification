# TCGPlayer-MTG-Picklist-Set-Code-Modification

TO USE THE PROGRAM:
1- Download the executable
2- Download a picklist from your TCGPlayer sale
3- Drag the CSV file over the executable. 
4- It will create a new XLSX file called Modified Pullsheet with the changes
5- It (should) differentiate between different card games so it will only make these changes for Magic The Gathering. So if you have sales for Yugioh or Pokemon it wont change those
6- The sets it recognizes is based on the INI file it generates. You can add to the INI file to add sets for it to modify. The code by default only checks 1 set. I'll likely update the project in the future to add more sets to the default INI but in the mean time you can add them yourself.

Takes the default TCGPlayer CSV picklist file and checks the set name against an internal INI to add a column with the set codes. This can help with inventory for medium to large sellers if you want to organize sets by the set code (where available) and easily match set codes on pick lists so you don't need to memorize set icons.

This project is mostly made for personal use and idk how github or pull requests or anything like that works. If you wanna use it, go ahead. I can't help if you have issues though because I myself don't know what I'm doing. 

###
Example usage:
Source file (directly downloaded from TCGPlayer):
Product Line	Product Name	Condition	Number	Set	Rarity	Quantity	Main Photo URL	Set Release Date
Magic	Titan's Strength	Near Mint	166	Magic Origins	C	1		7/17/2015 0:00
Orders Contained in Pull Sheet:	[Redacted]

Final output:
Product Line	Product Name	Condition	Number	Set	Rarity	Quantity	Main Photo URL	Set Release Date	Set Code
Magic	Titan's Strength	Near Mint	166	Magic Origins	C	1		7/17/2015 12:00:00 AM	ORI
###

Notice how it added a new column, "Set Code", and added "ORI" at J2 to signify the set code for this sale. The set codes can be found on the bottom left for most modern cards, usually 3 letters, so organizing inventory by the set code can be helpful and this makes it so you don't need to guess what the set code is.


