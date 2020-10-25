# Invoice-Generator
My father and brother are both self-employed truck drivers. They both receive a list of charter costs every week in a PDF file in different formats. This program reads the text in the PDF file and searches for patterns in the text. With the GUI (image below), I can select which company I'm generating the invoice for. This way the program knows the format of the list of charter costs. It also fills in the correct company data in the Word template (second image below). 

After filling in the the company data, the program generates a table in the Word template, filled in with the results of the pattern search. 

The program GUI. It automatically generates an invoice number based on the company, and the current year and week. It only takes a couple of clicks to generate an invoice in Word format and save it in PDF format as well:

![alt text](https://user-images.githubusercontent.com/58829624/97108797-0c483700-16d0-11eb-980d-88b519568926.png)

The empty Word template that gets filled in:

![alt text](https://user-images.githubusercontent.com/58829624/97108795-0bafa080-16d0-11eb-8e19-2602c50737f3.png)

With the GUI, it's also possible to select up to 100 different table styles. Examples of generated invoices below (with sensitive information removed): 

Example 1:

![alt text](https://user-images.githubusercontent.com/58829624/97109168-f0459500-16d1-11eb-8bfd-da74982364ad.png)

Example 2:

![alt text](https://user-images.githubusercontent.com/58829624/97109170-f0de2b80-16d1-11eb-82a9-a5f9e538894b.png)

Future plans are to clean up the code a bit and deploy this program to a web application.
