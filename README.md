# This repository has been archived.

# MVC-Excel-Extensions

<p>Allows developer to return an <em>ExcelResult</em> by passing a collection of any type to <em>Excel</em> method just like <em>JsonResult</em> and <em>Json</em> method, Just call the method, pass the collection you want its data to be exported to excel file and watch the magic.</p>
<p>Also you can mark some properties to be ignored by <em>ExcelResult</em> while generating the Excel file by setting <em>ExcelIgnoreAttribute</em> on the property you want to ignore and more..</p>

<h3>Setup your controller</h3>
<p>In order to have <em>ExcelResult</em> to work you need to specify Excel method in your base controller if you have one, If you don't just let your controller inherit from <em>ExcelMasterController</em></p>
<h4>1 - Inherit From <em>ExcelMasterController</em>:</h4>
<pre>
<code>
    public class HomeController : ExcelMasterController
    {
    
    public ActionResult ExportToExcel()
    {
        List&lt;Person&gt; lst = new List&lt;Person&gt;();
        for (int c = 0; c < 50; c++)
        {
            lst.Add(new Person() {BirthDate = new DateTime(1990, 12, 18), IsMale = true, Name = "Ibrahim", Summary = "I'm passionate about technology specially Microsoft's technologies in software development, Adore software development and willing to be one of the worldwide noted developers, getting the most of my technical skills along with my studies in business field to develop a real business solutions.I'm a knowledge hunger have no end point for my learning path and trying to reach for the sky." });
        }
        return Excel(lst, "TEST REPORT","TEST SHEET");
    }
    }
</code>
</pre>
<h4>2 - Defining your own Master Controller:</h4>
<pre>
<code>
     public class ExcelMasterController : Controller
    {
        protected ExcelResult&lt;T&gt; Excel&lt;T&gt;(IEnumerable&lt;T&gt; data) where T : class
        {
            return new ExcelResult&lt;T&gt;(data);
        }
        protected ExcelResult&lt;T&gt; Excel&lt;T&gt;(IEnumerable&lt;T&gt; data, string fileName) where T : class
        {
            return new ExcelResult&lt;T&gt;(data, fileName);
        }
        protected ExcelResult&lt;T&gt; Excel&lt;T&gt;(IEnumerable&lt;T&gt; data, string fileName, string sheetName) where T : class
        {
            return new ExcelResult&lt;T&gt;(data, fileName, sheetName);
        }
   }
</code>
</pre>

<p>
    <em>ExcelResult</em> Have several overloads that takes some parameters described as following
</p>
<table class="table table-striped">
    <thead>
    <tr>
        <th>Parameter</th>
        <th>Type</th>
        <th>Description</th>
    </tr>
    </thead>
    <tbody>
    <tr>
        <td>data</td>
        <td>IEnumerable&lt;T&gt; where T : class</td>
        <td>The data you want to be exported to Excel File, It takes a generic collection but the generic type is constrained to be a class.</td>
    </tr>
    <tr>
        <td>fileName (Optional)</td>
        <td>String</td>
        <td>The name of the file will be downloaded by the client (Default will be the name of the type passed to the generic data collection).</td>
    </tr>
    <tr>
        <td>sheetName (Optional)</td>
        <td>String</td>
        <td>The name of the sheet in the excel file (Default will be the name of the type passed to the generic data collection).</td>
    </tr>
    </tbody>
</table>
<br />
<hr/>

<h3>Format and Customize your ExcelResult:</h3>
<p>MvcExcelExtensions contains some attributes that can be used to format your returned excel file, Just mark your class's Properties with the attributes and it will work.</p>
<p>Following is <em>Person</em> class which i used above to generate <em>ExcelResult</em> by passed a list of it to <em>Excel&lt;T&gt;</em> method.</p>
<pre>
<code>
    [ExcelSheetStyle(HeaderFontSize = 14, BodyFontSize = 12, FontFamily = "Times New Roman", IsHeaderBold = true)]
    public class Person
    {
        public string Name { get; set; }
        [ExcelValueFormat("dd, MMMM yyyy"), ExcelDisplayName("Birthday")]
        public DateTime BirthDate { get; set; }
        [ExcelIgnore]
        public bool IsMale { get; set; }
        [ExcelColumnStyle(HorizontalAlignment = HorizontalAlign.Left, VerticalAlignment = VerticalAlign.Top, Width = 35, WordWrap = true)]
        public string Summary { get; set; }
    }
</code>
</pre>

<br />
<hr/>

<h4>ExcelSheetStyleAttribute</h4>
<p>You can mark your class with this attribute to specify some global settings you want in your resulting Excel file.</p>

<table class="table table-striped">
    <thead>
    <tr>
        <th>Parameter</th>
        <th>Type</th>
        <th>Description</th>
    </tr>
    </thead>
    <tbody>
    <tr>
        <td>HeaderFontSize (Optional)</td>
        <td>float</td>
        <td>Set the header's font size of the resulting excel file where header is the first row. (Default 14)</td>
    </tr>
    <tr>
        <td>BodyFontSize (Optional)</td>
        <td>float</td>
        <td>Set the body's font size of the resulting excel file where body is all rows except the first. (Default 12)</td>
    </tr>
    <tr>
        <td>FontFamily</td>
        <td>String</td>
        <td>Set the font family that will be used for the text inside the resulting excel file. (Default 'Times New Roman')</td>
    </tr>
    <tr>
        <td>IsHeaderBold</td>
        <td>bool</td>
        <td>Specify if you want the header (i.e, First Row) to be bolded. (Default true)</td>
    </tr>
    </tbody>
</table>

<br />
<hr/>

<h4>ExcelValueFormatAttribute</h4>
<p>You can set a format for a specific property to be used in the resulting excel file. Note that the type of the property marked with this attribute must implement <em>IFormattable</em> interface otherwise the format will be ignored.
</p>
<p>This attribute takes only string representing the format in its constructor and its mandatory</p>

<br />
<hr/>

<h4>ExcelDisplayNameAttribute</h4>
<p>Use this attribute if you don't want the resulting excel file's header contain the property name but contains another name, In the example above i decided to use 'Birthday' as a header for 'BirthDate' values</p>
<p>This attribute takes only string representing the header title in its constructor and its mandatory</p>

<br />
<hr/>

<h4>ExcelIgnoreAttribute</h4>
<p>Use this attribute if you don't want a property to be printed in the resulting excel file.</p>
<p>This attribute takes no arguments and have no parameters</p>

<br />
<hr />

<h4>ExcelColumnStyle</h4>
<p>Use this attribute to style the column that will list the property's values.</p>
<table class="table table-striped">
    <thead>
    <tr>
        <td>Parameter</td>
        <td>Type</td>
        <td>Description</td>
    </tr>
    </thead>
    <tbody>
    <tr>
        <td>HorizontalAlignment (Optional)</td>
        <td>HorizontalAlign</td>
        <td>Set horizontal alignment for all column's cells (Default HorizontalAlign.Center)</td>
    </tr>
    <tr>
        <td>VerticalAlignment</td>
        <td>VerticalAlign</td>
        <td>Set vertical alignment for all column's cells (Default VerticalAlign.Middle)</td>
    </tr>
    <tr>
        <td>Width (Optional)</td>
        <td>double</td>
        <td>Set a custom width to the column that will list the property's values. (Default To Fit Content)</td>
    </tr>
    <tr>
        <td>WordWrap</td>
        <td>bool</td>
        <td>Set to true if you want the word to be wrapped. (Default false)</td>
    </tr>
    </tbody>
</table>


<p>
    <i>Special Thanks To <strong>EPPlus</strong> Team, I've used their great library in CodePlex to create mine. So special thanks for them. You can find EPPlus library <a href="https://epplus.codeplex.com/">Here</a></i>
    
</p>
