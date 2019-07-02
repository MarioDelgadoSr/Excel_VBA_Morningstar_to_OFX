<!-- Markdown reference: https://guides.github.com/features/mastering-markdown/ -->

# *Excel_VBA_Morningstar_to_OFX*

* This VBA program will generate an OFX file from [Morningstar](https://www.morningstar.com/)'s Portfolio export and into an [OFX formatted file](http://moneymvps.org/faq/article/8.aspx).  
* The OFX file can then be imported into [Microsoft Money Plus Sunset](https://www.microsoft.com/en-us/download/details.aspx?id=20738) to update the portfolio's stock and mutual fund prices.

With this VBA program installed in Excel, you have a reliable, free source of stock and mutual fund data to keep your Microsoft Money portfolio upto date.

## Background: [Obtain stock and fund quotes after July 2013](http://moneymvps.org/faq/article/651.aspx)

## Instructions

* Add [Excel_VBA_Morningstar_to_OFX.vba](https://github.com/MarioDelgadoSr/Excel_VBA_Morningstar_to_OFX/blob/master/vba/Excel_VBA_Morningstar_to_OFX.vba) to Excel.

	* **Related**:
	
		* [How to insert and run VBA code in Excel - tutorial for beginners](https://www.ablebits.com/office-addins-blog/2013/12/06/add-run-vba-macro-excel/)
		* [Copy your macros to a Personal Macro Workbook](https://support.office.com/en-us/article/Copy-your-macros-to-a-Personal-Macro-Workbook-AA439B90-F836-4381-97F0-6E4C3F5EE566)
		
* **Edit the program** (line 53) to specifiy the location of the OFX file that will be generated. By default, the program will write out to "C:\temp\quotes.ofx".

* Create a portfolio in Morningstar with 'Ticker', '$ Current Price' and 'Morningstar Rating For Funds' columns.  

	* [Video: Creating a Morningstar Portfolio](http://video.morningstar.com/us/misc/portfoliomanager/portfolio_noexisting.html)
	* ![Screen Shot of required column in custom portfolio view](https://github.com/MarioDelgadoSr/Excel_VBA_Morningstar_to_OFX/blob/master/img/portfolio.png)

* Use the Morningstar 'Export' utility to export the custom portfolio view to Excel.

* Run the *makeOFX_file* macro in Excel.  It will dynamically read the Excel-based portfolio data created by Moningstar and write out the specified file.

* Installation of Microsoft Money should assoicate .ofx file with the Microsoft Money Import Handler [mnyimprt.ext](http://moneymvps.org/faq/article/407.aspx).  
  Double clicking your file will in the file explorer will start the import handler and prompt you to start Money to continue with the import.
  
	* To automate the import, you can create a [desktop shortcut](https://answers.microsoft.com/en-us/windows/forum/windows_10-start/quick-tip-create-desktop-shortcuts-in-windows-10/d867565e-34c2-42ad-88da-ccf76a4a9820) for the quotes.ofx file.
  
* View your updated Microsoft Money Portfolio.


## Author

* **Mario Delgado**  Github: [MarioDelgadoSr](https://github.com/MarioDelgadoSr)
* LinkedIn: [Mario Delgado](https://www.linkedin.com/in/mario-delgado-5b6195155/)
* [My Data Visualizer](http://MyDataVisualizer.com/demo/): A data visualization application using the *DataVisual* design pattern.


## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details




