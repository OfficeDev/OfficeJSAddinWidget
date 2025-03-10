# Sample code for the Web Widget that facilitate Office add-in distribution.

## Introduction
This project simplifies the process of distributing Office add-ins by enabling direct installation from a website. We provide a convenient install link and a web widget to enhance user experience and streamline the add-in installation process for better acquisition conversion.

## Getting Started

The webpage will show the add-in name, a few details about the add-in functionality for highlight, a dropdown button to install and use the add-in and an image to demo the add-in. This is what the UI looks like.

<img alt="DemoUI.png" src="https://github.com/OfficeDev/OfficeJSAddinWidget/blob/main/Example-ScriptLab-UI.png">

> [!IMPORTANT]  
> To use this sample code, please follow the [Partner-Led Marketing Guidelines](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fforms.office.com%2Fpages%2Fresponsepage.aspx%3Fid%3Dv4j5cvGGr0GRqy180BHbR9SZm6iKPzNJvudw-PPFJydUNFZOMTJVWlVaOTRHQVQ5RENQMUEwWVdaMC4u&data=05%7C02%7Camha%40microsoft.com%7C65489391094a4ceb72fe08dc488e5159%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C638465023522998462%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C0%7C%7C%7C&sdata=vVQh3Binli51i8S8T81j%2F4b9%2F0OtDfXYUj1q2HWug1g%3D&reserved=0).

> [!TIP]
> If you want a simple trial on the dropdown button link that enables user to install and use add-in, please clone https://github.com/OfficeDev/OfficeJSAddinWidget/blob/main/LinkGenerator.html to your local, open with browser app and input the information as the UI shows.

### Usage for how to generate the install link for Office on Windows or Mac

To better distribute your add-in, you can create an install link to provide your users with the "click and run" experience from your website or anywhere else after you submit your add-in to AppSource. The link will seamlessly bring users to their Web sersion WXP with your add-in automatically launched, so you can directly guide users to try your add-in instead of letting them find it themselves in the Add-in Store. 

To create the link, you can reference the link below and configure the parameters in the query string.<br>
<a href="url">{{appName}}:https://api.addins.store.office.com/addinstemplate/{{language}}/{{correlation}}/{{addinId}}/none/{{addinName}}.{{fileFormat}}?omexsrctype=1&isexternallink=1</a>

1. <strong>appName</strong><br>
   This parameter controls which Office application would be opened when users click the link. <br>
	· For Word: ms-word <br>
	· For Excel: ms-excel <br>
	· For PowerPoint: ms-powerpoint <br>
	This approach currently supports WXP on Windows and Mac only.<br>

> [!Note]
> This parameter is part of the deeplink.

2. <strong>language</strong><br>
   This is the language of Add-in instructions after opened on Word/Excel/PowerPoint on Windows or Mac.<br>
   Please set as needed. For example: "de-DE", "fr-FR", "it-IT", "ja-JP", "zh-CN".<br>

3. <strong>correlation</strong><br>
   A GUID for diagnostic purpose. For example, "7bf846ec-905a-5edd-b162-83498f9a8674".<br>
   We strongly recommend you to runtime generate it and make it different per click.<br>

4. <strong>addInId</strong><br>
   The ID of your add-in published in AppSource<br>

> [!NOTE]
> This parameter is case sensitive. Please use upper case in order to succeed.

5. <strong>addInName</strong><br>
   The full name of your add-in in percent encoding format. For example, "space" should be %20, and "-" should be %2D...etc.<br>

6. <strong>fileFormat</strong><br>
    · For Word: docx <br>
	· For Excel: xlsx <br>
	· For PowerPoint: pptx <br>

For example, if we want to create an install link so that users can click and launch Script Lab on Word on Windows, this link should be:
<a href="url" target="_blank">ms-word:https://api.addins.store.office.com/addinstemplate/en-US/228a829b-69d7-45f4-a338-c6aba330ec7e/WA104380862/none/Script-Lab--a-Microsoft-Garage-project.docx?omexsrctype=1&isexternallink=1</a>.<br>

### Usage for how to generate the install link for Office on the Web

To better distribute your add-in, you can create an install link to provide your users with the "click and run" experience from your website or anywhere else after you submit your add-in to AppSource. The link will seamlessly bring users to their Web sersion WXP with your add-in automatically launched, so you can directly guide users to try your add-in instead of letting them find it themselves in the Add-in Store. 

To create the link, you can reference the link below and configure the 3 parameters in the query string.<br>
https://go.microsoft.com/fwlink/?linkid={{linkId}}&templateid={{addInId}}&templatetitle={{addInName}}

1. <strong>linkId</strong><br>
   This prarmeter controls which web endpoint would be opened when users click the link. <br>
	· For Word Web: 2261098 <br>
	· For Excel Web: 2261819 <br>
	· For PowerPoint Web: 2261820 <br>
   This approach currently supports WXP Web only.<br>

2. <strong>addInId</strong><br>
   The ID of your add-in published in AppSource<br>

3. <strong>addInName</strong><br>
   The full name of your add-in in percent encoding format. For example, "space" should be %20, and "-" should be %2D...etc.<br>
   The name should not be longer than 128 characters.

For example, for the add-in Script Lab, it supports all Word, Excel, PowerPoint. 
If we want to create an install link so that users can click and launch Script Lab, this link should be:
https://go.microsoft.com/fwlink/?linkid=2261819&templateid=WA104380862&templatetitle=Script%20Lab,%20a%20Microsoft%20Garage%20project
Where 2261819 is the Excel Web launch link ID, WA104380862 is Script Lab's add-in id, and "Script%20Lab,%20a%20Microsoft%20Garage%20project" is the full name in percent encoding format.<br>

### Usage for the web widget

1. Download the WebWidget/Example-ScriptLab.html file and go to <script> at line 62.

2. Config the paramenters under "Paramenters that need to config" part.<br>
	a. <strong>addinId</strong><br>
		This is the unique add-in ID. You can get the correct value by following below steps.<br>
		&emsp;1) Go to https://appsource.microsoft.com/en-US/ from your browser.<br>
		&emsp;2) Input your Office add-in name in the search bar on top center of AppSource homepage.<br>
		&emsp;3) Click your add-in in the seach results.<br>
		&emsp;4) The add-in information page will be automatically displayed in current tab.<br>
		&emsp;5) The add-in ID is in the URL.<br>
   For example, if the URL is https://appsource.microsoft.com/en-US/product/office/WA104380862?tab=Overview, then "WA104380862" is the add-in ID that you should input for this parameter in sample code.
		
	b. <strong>addinName</strong><br>
		This is the add-in name. You can get the correct value by following below steps.<br>
		&emsp;1) Go to the webpage in 2.a.4).<br>
		&emsp;2) The add-in name is displayed as the title on right of the add-in icon.<br>
   For example, if the webpage is https://appsource.microsoft.com/en-US/product/office/WA104380862?tab=Overview, then "Script Lab, a Microsoft Garage project" is the add-in ID that you should input for this parameter in sample code.
		
	c. <strong>wordOnlineSupported, excelOnlineSupported, powerpointOnlineSupported, wordDesktopSupported, excelDesktopSupported, powerpointDesktopSupported</strong><br>
		This is the Office products that this add-in supports. You can get the correct value by following below steps.<br>
		&emsp;1) Go to the webpage in 2.a.4).<br>
		&emsp;2) Click "details + support" tab on the webpage.<br>
		&emsp;3) Scroll down to "Products supported" section.<br>
			&emsp;&emsp;- If "Word on the web" is in the list, then set wordOnlineSupported to true. Otherwise, set it to false.<br>
			&emsp;&emsp;- If "Excel on the web" is in the list, then set excelOnlineSupported to true. Otherwise, set it to false.<br>
			&emsp;&emsp;- If "PowerPoint on the web" is in the list, then set powerpointOnlineSupported to true. Otherwise, set it to false.<br>			
			&emsp;&emsp;- If "Word on Windows" or "Word on Mac" is in the list, then set wordDesktopSupported to true. Otherwise, set it to false.<br>		
			&emsp;&emsp;- If "Excel on Windows" or "Excel on Mac" is in the list, then set excelDesktopSupported to true. Otherwise, set it to false.<br>		
			&emsp;&emsp;- If "PowerPoint on Windows" or "PowerPoint on Mac" is in the list, then set powerpointDesktopSupported to true. Otherwise, set it to false.<br>
	
	d. <strong>language</strong><br>
		The language of the Add-in instructions after opened on Word/Excel/PowerPoint on Windows or Mac. For example, "en-US", "de-DE", "fr-FR", "it-IT", "ja-JP", "zh-CN". Please leave this blank if your Add-in is only available on Office on the Web.

	e. <strong>linkAddInAppSource</strong><br>
		If you would like to show the link to Office AppSource, then set linkAddInAppSource to true. Otherwise, set it to false.

	f. <strong>addinDetails</strong><br>
		This is for descriptions about the add-in functionalities displayed on the webpage. 
		
	g. <strong>demoImage</strong><br>
		This is the image for the add-in that display on the webpage.
		
4. Save the html file and open by browser. Verify the UI and dropdown links works for your scenario.

5. Make any additional changes to the sample code as you need, and integrete it into your website.

## Special Notice

The deeplink can successfully open an Office document with your add-in and the end users can use your add-in for below scenarios only.

1. The add-in is public published, so that it can be found from Office AppSource https://appsource.microsoft.com/.

2. The Office store is enabled for the end user.




## Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.opensource.microsoft.com.

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft 
trademarks or logos is subject to and must follow 
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.
