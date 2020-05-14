# Excel-Data-Cleaner

<h3> Problem </h3>
While doing sales prospecting, I noticed I was spending a lot of time copying and pasting individual names and contacts into Microsoft Excel. I needed a way to reduce the amount of time I spent manually entering data into Excel.


<h3> Solution </h3>

I created an Excel plug-in that analyzes a spreadsheet containing all of the unstructured sales data and correctly pulls out the "Name", "Email","Company Name", and "Title" for each individual prospect. Instead of having to switch between screens to copy and paste four times per prospect (one copy and paste per data point), I only have to do so no more than one time per prospect now. 

<h3>Some Implementation Details</h3>


To create the script, I first had to initialize Office-Add-ins as explained <a href="https://docs.microsoft.com/en-us/office/dev/add-ins/tutorials/excel-tutorial"> here</a>.

I reviewed my data source and determined that for each prospect, there are 24 data fields available. So, to use the tool I would either copy all 24 data points at once for each prospect from my data source or copy multiple groups of 24 at the same time.

Once the script is run, the 4 key data points are extracted from each batch of 24 and then reinserted as 4 cells in a new row in the spreadsheet.



