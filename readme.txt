Smart Storage with 2 compressions(zip/huffman)

Smart Storage is something like a file packager, you can add any file(s) of any format to the storage file(AddUpdateItem function) and remove them(DeleteItem function), of course. 

Now it comes into two separate version: with/without compression.

You can choose zlib/huffman or no compression for each file.

It uses index(file) technique for a better performance, thus it has VaccumStorage() function. 

And, the most inportant part is that it uses serialized section(chunk) technique to handle the CRC & compression of large file.

And this project is originally aimed to be the storage part of Carles P.V's Thumbnailer 1.0 (image thumbnailer-viewer with GDI+)(http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=59677&lngWId=1), thus it includes the PictureFromByteStream() function to directly return a picture object from a byte array! 

Beside that, there are: GetItemText() function for direct return of plain text content; GetItemPicture() function for direct return of picture object according to the index; SaveItemToFile() function to save the content to disk file according to the index.

Some of the codes are not written by me, such as cHuffman, cCRC. They are from psc, now they are back for you guys:)

The zlib dll is generated at runtime of sample test since psc will remove all PE file. It is easy for you to remove the zlib dependancy, that's why I only enclosed it in the sample project.

And I make all these functions into a class for handy usage. I also use Copymemory for a better performance. Please feel free to leave any comments, bugs or suggestions are welcome!

Usage: 
1.Simply open the sample storage, then click listed files, content displayed, doubleclick to save.
2.Select DB File->Create->Open->Select Item File, add it...

Sorry for lack of code comments, but I think that the method is really simple.
