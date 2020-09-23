Hi there!

This is my 2nd Upload to the PSC.  I know PSC a several years and I should say that this Site isnt only a Internet-Site like the most -  PSC is better be called as the "Codeportal" !. Enough Promotion! 
When i was looking for a cool code to store data on my SQL Server Database i found Pete Sral's Code what's posted in the PSC. Thanks a lot Pete! Now that I need to have a Frontend what is able to store big Blobs of Data to a Access Database I tried out the AppendChunk and GetChunk - Method - but i have trouble to restore big blobs to a file again!. I remember about Pete's Code of
HOWTO: Access and Modify SQL Server BLOB Data by Using the ADO Stream Object - originally based of the MS Article ID: Q258038.

I modified this code again.

1. Adding Registry-Access Functions 
2. Adding a Module to create (G)lobal(U)nique(I)(D)entifiers GUID's 
3. Adding a Class to Connect to Databases
4. A New Access 2000 Database called Blob.mdb, containing the Table Personal from the GermanVersion of the Northwind Database
5. last but not l... - a new Form.


Make sure:
-a Folder named "Temp"exists in the Application-Path
-the Blob.mdb -file exists in the Application-Path

If you want to use a function that's able to store BIG BAD BLOBS don't mess around with the Chunks - Stream it!

Please feel free to:

- use/modify this code for your own needs. 
- distribute this code in you AppZ to others (dont forget to give a credit to me ;-)
- mail me if you have any comments or questions


marcuslauermann@gmx.net

IMPORTANT:

IF YOU MAKE ANY MODIFICATION TO A FINAL BUILD - DO NOT HIDE INFORMATION - SHARE YOUR CODE TO OTHERS CAUZ' THIS IS FAIR TO ALL OTHA'Z




