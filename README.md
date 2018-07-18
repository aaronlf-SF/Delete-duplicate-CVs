# Delete-duplicate-CVs

In theory, the process of uploading CVs to Caliber should be quick and easy. In practice, it's not.

Before CVs can be uploaded, they first have to be manually checked. This is due to the fact that there are often many duplicates, CVs without email addresses, and documents that aren't CVs at all. This is an unavoidable inconvenience, simply a result of the way CVs are exported in bulk from consultants' email folders.

With the help of this script, the process of uploading CVs can be swift and simple, as it should be. 

By specifying 'PATH' to the directory of the CVs, duplicate CVs can be detected and deleted all with the press of a button.

At time of writing, this script only supports .docx and .pdf file formats. ~~It is hoped that .odt and .rtf formats will be possible to implement also.~~ The formats .odt and .rtf will not be implemented by this software. This is because the number of files is so low that it would not be worth the time investment. 

It appears that .doc files (1997-2003 Word documents) are now possible to implement! (13/07/2018)


### Extra additions made:

#### (16/07/2018)
When deleting duplicates, CVs are organized by date and the most recent one is kept, instead of doing it randomly


#### (18/07/2018)
All files within a folder containing the same email address will be compared and all but one will get deleted. This is an innovation far superior to the previous script, which could only compare two documents that were one next to another in alphabetical order.

