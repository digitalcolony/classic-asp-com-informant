classic-asp-com-informant
=========================

Tool to test what 3rd party components are installed on the web server. 

(November 9, 2001)

I currently do business with 4 different web hosts. Each of them has a different subset of 3rd-party ASP components installed on their server. Sometimes they are open with which components are installed, sometimes they aren’t. Whenever I wanted to test to see if a particular COM object was available, I’d write a quick script. The script would try to create the object via the Server.CreateObject method and then I’d go to the page to see if it returned an error code. No error code meant it was installed and I could start coding my application around that knowledge.

Tedious and Repetitive

After about the 10th script I wrote, it hit me that there probably is a better way to do this. What was needed was a script that tested the most common ASP components and allowed the user to quickly add new ones to the list. Having a bunch of test scripts lying on the file server wasn’t optimal. And the last thing you would want is to hard-code all the test cases inside your ASP code. My solution was to have a single page handle the creation, modification, and display of component test cases. The data source would be a single XML file.

Developers Tool

COM Informant is handy tool to have if your development team creates custom components and then deploys them across multiple servers. What better way to test if a component is installed than viewing a single web page. The top portion of the tool allows the user to add any component name to test list.
