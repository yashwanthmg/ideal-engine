While working on the size chart image expansion I had difficulty in completing the task on time as the Image file naming guidelines was a resource intensive task. The guidelines states that 
“Image files must be named by ASIN + variant code + file extension (for example, B000123456.MAIN.jpg). If uploading multiple images, you must create a ZIP file with all images inside before uploading.”
The problem is vendors name their files with SL No. , SKU or a similar product identifiers. To rename each file according to the above mentioned criteria, vendor has to first create a list of names on excel and then rename the file by copy paste. Due to fatigue of doing the same repetitive task there are chances of errors which is going to cost more time.

When looking for alternatives on amazon seller central & vendor central I did not come across any bulk image rename tool although there are tools for different purposes like Listing uploader, Category specific Product template, Price & quality template etc. So I had to find a way to get this done quicker. Hence I created a bulk image rename tool using excel VBA macros with the help of internet resources. Using this tool you can rename images or files and zip it with few clicks without having to rename each file manually. 

Right now I have created basic working version: Bulk Rename v1 (Single file) tool and its production ready since I have already used this tool for Size chart image expansion along with Bulk Rename v1 (Multiple files). But the latter requires few more modifications like SKU-ASIN VLOOKUP for auto fill and former need to be a bit more efficient like creating zip file with a click instead of popup which asks to select folder (again) This will be completed once I come back from my holiday. 

Below is the comparison of approx. time required for Manual vs Bulk Rename v1 (Single file) tool which was tested with 1 ASIN & 10 product images(Main, PT01, PT02…PT09) 

 

So as per calculations to rename 100 ASINs it will take 0.30 X 100 = 50 mins Approx. 
The Bulk Rename v1 (Multiple files) tool once modified and updated will be able to bring down this number to 10 mins for any number of ASINs.

Possible Scenarios for the need of this tool
1.	Want to update correct images if wrong images were uploaded due to human error.
2.	Want to replace existing images with new images.
3.	Want to include specific image as in the case of size chart expansion activity to comply with amazon guidelines. And so on..

