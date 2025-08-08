# Alma Printing With Powershell
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
### Preamble
In recent years, Ex Libris — the Library systems supplier of the [Alma Library Management System](https://exlibrisgroup.com/products/alma-library-services-platform/) — have made significant improvements to the Alma printing stack. Alma is a web-based, cloud-hosted solution, that historically relied upon customers' own email infrastructure for printing. Print jobs would be sent by email to a nominated customer Inbox representing the print queue. It would then be up to the customer to process and print the incoming emails, and while this works, this method of printing requires some effort on the part of the customer to implement a solution to this side of the workflow. Such solutions are often fragile and prone to failure.

A few years back, Ex Libris introduced two new methods of printing in Alma in addition to email printing:

The first method is called "Quick printing". When enabled by the operator, the Alma web app passes the print data to the browser, and the browser receives a signal to invoke the browser print dialogue, enabling the operator to send the job to their preferred printer. This is useful for occasional ad hoc printing, and is assumed to have been made possible because of the availability of new printing APIs in modern web browsers. However, this is not useful for printing scenarios where a workflow is involved, because multiple clicks are required, which would soon add up to an increase in time spent by the operator processing the items.

The second method, the Printouts Queue, addresses this key drawback of Quick printing. The HTML print content can now be fetched from the Printouts Queue using a new Alma Printing API, and [Ex Libris's Alma Print Daemon](https://github.com/ExLibrisGroup/alma-print-daemon) provides a client solution for customers to implement this method.

However, Ex Libris's Alma Print Daemon appears to have some downsides:
- being an Electron app it is very large; the installer is over 100MB!
- it is built on an old version of `node.js`, and I have tried & failed to repackage with `node.js` 16, due to a problem with the dependency `node-native-printer`
- it only supports standard paper sizes. We want to be able to print to `Roll Paper 80 x 297 mm` on  Epson TM Series POS printers, and attempts to do this with the daemon results in very small-font printouts

The Powershell script provided in this repository is an attempt to address the above downsides, by offering an alternative "client", although in its current state it has a few downsides of its own, which are explained later on.

### What does the script do?

The Powershell script contained in this repository polls the Alma print queue for new jobs. If it finds any it will send the job to the specified printer and mark the job as `Printed`.

![An image of the script in operation](./images/alma_printing_screenshot.png?raw=true)

The intention is that the script runs in the background, with its window minimised to the taskbar by default. If something goes wrong then the taskbar icon can be clicked to show the above window.

![An image of the taskbar icon](./images/alma_printing_taskbar_screenshot.png?raw=true)

### How does the script work?

The script works by running a loop to check the queue every `X` seconds. The number of seconds can be changed with the `-checkInterval` parameter of the `Fetch-Jobs` function. By default, this is `30` seconds. The communications with Alma's API is achieved with Powershell's `Invoke-RestMethod`.

The HTML content is saved to a file stored in the `tmp_printouts` subdirectory. Each file contains the content for one printout, and the filename will contain the `letterId` (the ID assigned by Alma to the printout). If something related to printing goes wrong, but the "fetch" from Alma succeeded, it's possible to manually print the content from the stored file, e.g. by opening this file in a web browser and <kbd>Ctrl</kbd> + <kbd>P</kbd>.

Rendering and printing the HTML is achieved using an Internet Explorer COM object. This is probably the weakest area of the script; there is currently no way to internally specify which printer to print to, so the script relies upon setting the printer specified by `-localPrinterName` as the Windows default printer when there is a job to print, restoring the default printer afterwards. The script also does something quite similar for the `Page Setup` settings (margins, etc), reading in the existing settings, changing them when there are jobs to print, and restoring the original settings when done. Though these settings are not printer-specific, typically we'll want different values according to the printer we want to print to.

By default, the script will only fetch printouts that have the Alma printout status `Pending`. Once the printout has been printed, the script attempts to change the status to `Printed`. There is a `Fetch-Jobs` function parameter `-printoutsWithStatus` that can be added, which can be used to fetch printouts with other statuses (`Printed`, `Pending`, `Canceled`, `ALL`). For example, `-printoutsWithStatus 'Canceled'`

### Prerequisites
There are tasks that need to be done in the Alma web UI Configuration area, so that the Alma printer is available as an online queue, which are [explained here](https://knowledge.exlibrisgroup.com/Alma/Product_Documentation/010Alma_Online_Help_(English)/030Fulfillment/080Configuring_Fulfillment/020Fulfillment_Infrastructure/Configuring_Printers). For the purposes of running this script you will need to [get an API key](https://developers.exlibrisgroup.com/blog/how-to-set-up-and-use-the-alma-print-daemon/) to use.

### How should I deploy the script for non-technical staff to use?

Clone this repository to the machine you want to run it on, and open an elevated Powershell window.

Once you have a suitable Alma API key, you'll want to make it available for the script to use:
```
Set-Location <script-dir>
. .\FetchAlmaPrint.ps1;Invoke-Setup
```
After running the above commands, you will be prompted to enter your API key, which you can paste in. This is then stored in an XML file for future use by the script.

> For UoY deployments, skip the next two steps and jump to [Setup steps for starting Fetch-Jobs automatically](#setup-steps-for-starting-fetch-jobs-automatically)

The second thing you'll want to do is determine the Alma print queue to monitor:

```
 . .\FetchAlmaPrint.ps1;Fetch-Printers
 ```
 This will list all of the Alma printers. You'll want to note the Printer ID for the queue you're interested in printing from.

 (note that you only need to dot source the script if it's no longer in memory, i.e. you closed the Powershell window between invocations)

 Thirdly, and lastly, you'll want to run the `Fetch-Jobs` function, with the parameters set according to the requirements, e.g.:
```
 . .\FetchAlmaPrint.ps1;Fetch-Jobs -printerId "848838010001381" -localPrinterName "EPSON TM-T88III Receipt" -checkInterval 20
```
In a production environment, it is the `Fetch-Jobs` function that will run in the background, polling for and printing new print jobs as they are generated at the back-end.

### Setup steps for starting Fetch-Jobs automatically

What follows is the setup required to start `Fetch-Jobs` automatically upon logon, and a desktop shortcut for the operator to re-launch the script if something untoward happens and the running instance needs to be restarted.

In a production environment, it is recommended that two identical script shortcut (`.lnk`) files be created. These should be put in the `shell:startup` (`C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp`) and `shell:desktop` (`C:\Users\Public\Desktop`) directories. The shortcut in `shell:startup` ensures the script runs automatically upon logon, and the shortcut in `shell:desktop` provides an option to manually invoke the script, if the Powershell window was accidentally closed. The `.lnk` files will ensure that the script will run minimised, so as to run discretely in the background, lessening the chance of an operator accidentally closing the window.

The script `DeployShortcuts.ps1` is available in the `setup-scripts` subdirectory, to make it easier to create and deploy the `.lnk` shortcut files. It is <ins>highly recommended</ins> to create the shortcuts using this script, because of File Explorer's 260 character limit in respect of viewing and editing shortcut properties' `Target` field. This limit will very likely be exceeded when including additional parameters, and in some environments with e.g. very long printer names, which need to be included as parameter values.

##### Running `DeployShortcuts.ps1`
> For a list of UoY-specific command lines for creating shortcuts, see the wiki page `FMSYS: Library - LMS printing`
1. Open an elevated Powershell window
2. Type `Set-Location 'C:\fmsys-alma-printing-api\setup-scripts'` (or the path to where the repo is, plus `\setup-scripts`)
3. Copy the following, editing the parameter values with `<>` placeholder characters, and adding any additional required parameters:
```
.\DeployShortcuts.ps1 -ShortcutFilename 'Alma Slip Printing' -ShortcutArguments "-NoLogo -NoProfile -Command `"& { Start-Sleep 30;. .\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId '<printerId>' -localPrinterName '<printerName>' }`""
```
4. Paste the resulting line into your Powershell window and press `CR` to run the script
<br><br>

<details>
<summary><h4 style="display:inline-block">Inspecting the properties of an existing `.lnk` shortcut file (click to expand)</h4></summary>

The `DeployShortcuts.ps1` script also allows for the inspection of the properties of existing `.lnk` shortcut files:

1. Open a Powershell window
2. Copy the following, editing the `-ListOnlyFilePath` parameter value if you want inspect a different shortcut:

```
.\DeployShortcuts.ps1 -ListOnlyFilePath "C:\Users\Public\Desktop\Alma Printing.lnk"
```
<hr>
</details>

<details>
<summary><h4 style="display:inline-block">"Legacy", manual deployment steps (click to expand)</h4></summary>
These are included here "<i>just in case</i>".<br>
To create the shortcuts manually, take the following steps:

1. Assuming this repository is in `C:\fmsys-alma-printing-api`, create a new shortcut in this directory:
   1. In File explorer, right-click in the whitespace, and from the menu left-click `New` > `Shortcut`
   2. In the resulting wizard, enter `powershell` for the location of the item. Click the `Next` button
   3. In the next screen, give the shortcut a suitable name like `Alma Slip Printing`. Click the `Finish` button


2. Now make some modifications to your new shortcut file:
   1. Right-click the shortcut file, and from the menu left-click `Properties`
   2. In the `Shortcut` tab of the resulting dialogue, edit the following fields:

Target:
```
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoLogo -NoProfile -Command "& { Start-Sleep 30;. .\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId '<printerId>' -localPrinterName '<printerName>' }"
```
Start in: `C:\fmsys-alma-printing-api`
Run: `Minimized`

3. Click the `OK` button to save these changes.
4. In the Windows `Run` app, enter `shell:common startup`. Explorer will open the startup directory.
5. Move your shortcut from `C:\fmsys-alma-printing-api` to this directory (_admin rights required_)
6. In the Windows `Run` app, enter `shell:common desktop`. Explorer will open the desktop directory
7. From `shell:common startup`, copy your shortcut to the `shell:common desktop` directory (_admin rights required_)
8. Lastly, log off the PC and on again to check that the script is starting normally.

<hr>
</details>

#### IE First-Run

Because the script relies upon Internet Explorer for rendering & printing the HTML, it is likely you'll see the following first-run box:

![An image of the Internet Explorer 11 first-run box](./images/IE11_First_Run_Image.png?raw=true)

To prevent this box from reappearing, a setup script is provided. Perform the following one-time step in an elevated Powershell window:
```
Set-Location 'C:\fmsys-alma-printing-api\setup-scripts'
.\DisableFirstRunIE.ps1
```

#### Housekeeping: `tmp_printouts`

The `tmp_printouts` directory will accumulate many HTML files over time. It's probably a useful contingency having the HTML files persisted to disk in case re-prints are required, or if there is a problem printing the document. However, over time these files may consume a significant amount of disk space, so it's recommended to delete or recycle the older ones while keeping the more recent ones. To help automate this housekeeping process, a Task Scheduler XML template and a Powershell script is provided to recycle files older than 30 days. This can be adjusted according to local needs; just edit the XML before importing, or modify the task once imported. See `maintenance-scripts\fmsys-alma-printing-api - clean tmp_printouts directory.xml` and `maintenance-scripts\recycleFiles.ps1`.

<sub>Note that `maintenance-scripts\silent.vbs` exists to run the Powershell script completely invisibly; `powershell.exe` has `-WindowStyle 'Hidden'` but this is only processed after the Powershell window has appeared.</sub>

#### Miscellaneous deployment tips

- Note the `Start-Sleep 30;` at the start of the `-Command` parameter. It was noted that if the script starts too quickly after logging in, then you might see errors. This ensures a delay before starting the queue checking.
- If you need to reduce the `font-size` for a particular printer, you can do this by adjusting the relevant XSL template in Alma. For example, we use the following to reduce the `font-size` when printed to our Epson POS receipt printers:
```
<xsl:if test="notification_data/receivers/receiver/printer/code = 'sorter' or notification_data/receivers/receiver/printer/code = 'kings'">
@media print {
  tr {font-size: 70%;}
}
</xsl:if>
```
The printer `code`s, referenced in the conditional statement, are returned when invoking the `Fetch-Printers` function.
- Sometimes when troubleshooting it might be useful to try manually printing from Internet Explorer to mimic what the script is doing. Unfortunately, as of June 2022, Internet Explorer, in its role as a web browser, [is no longer supported](https://techcommunity.microsoft.com/blog/windows-itpro-blog/internet-explorer-11-desktop-app-retirement-faq/2366549), with attempts to open Internet Explorer resulting in immediately being redirected to Microsoft Edge. However, we can work around this by constructing a line like as follows, and entering into Windows' `Run` box:
```
C:\Windows\System32\mshta.exe vbscript:Execute("Set oIE = CreateObject(""InternetExplorer.Application""):oIE.Visible = True:oIE.Navigate ""C:\fmsys-alma-printing-api\tmp_printouts\document-44945184140001381.html"", 0 : window.close")
```
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<sub>It's possible the above `mshta` invocation may trigger anti-virus alarm bells! YMMV!</sub>

### Barcode unreadable?
A problem was identified with the readability of the barcodes when printed using the script. The key points about this problem are:
- the printed barcodes appeared to be missing their right guard bar, which was causing the barcodes to be unreadable when scanned
- when the HTML content is rendered on screen in Internet Explorer, from the files stored in `tmp_printouts`, the guard bar is _not_ missing
- the Opticon branded scanners in use are incapable of reading the "faulty" printed barcodes, however, they can be read using ones' smartphone
- the problem is not present when the HTML content is printed from an alternative web browser such as Google Chrome
- the problem was narrowed down to it being Internet Explorer related, and there was [a very similar sounding problem described at this now-defunct link](https://social.technet.microsoft.com/Forums/windows/en-US/9276a5b1-24cf-4973-873c-768068617e79/issue-printing-with-internet-explorer-10-11?forum=ieitprocurrentver), which pinpointed the XPS subsystem (that IE uses for printing) as the root cause
- after experimentation, it was found that the barcodes could be converted from `PNG` to `JPG` format in order to resolve the readability problem

To provide a solution for this problem, a new `base64Png2Jpg` function was added which converts the base64 PNG data to base64 JPG. This can be used by adding the `Fetch-Jobs` function switch parameter `-jpgBarcode`.

Alternatively, another solution is to modify the XSL template associated with the printout so that — instead of including the barcode as an image — the barcode is included as text formatted with an installed barcode font such as [this one](https://www.idautomation.com/free-barcode-products/code39-font/). Example XSL snippet:
```
<p style="font-family: IDAutomationHC39M; font-size: medium;">(<xsl:value-of select="notification_data/phys_item_display/barcode"/>)</p>
```

### Future improvements

* Currently, if the script is interrupted while it is `Working..`, say by pressing `CTRL+C`, there's a chance that the original default printer and `Page Setup` settings as mentioned previously won't be restored. It might be possible to improve this by using `Try`,`Catch`,`Finally` [as indicated here](https://stackoverflow.com/a/15788979/1754517).
* The limitations of using Internet Explorer for printing could be overcome by using a third-party HTML rendering/printing tool [like this one](https://github.com/kendallb/PrintHtml). But sadly, having tested it, it doesn't cope with `Roll Paper 80 x 297 mm` paper size. Another option is the [Print HTML](http://www.printhtml.com/index.php) command line tool but it's old (~2009) and untested.
* It would also be good to see if this script could be made into a Windows service, perhaps using [NSSM](https://nssm.cc/), instead of invoking the script via shortcuts.
* To protect against a possible API endpoint security compromise, it would be a good idea to sanitise the HTML letter content before "opening" it, as is effectively done with `$ie.Navigate($printOut)`. The idea would be that this would protect against e.g. malicious `<script></script>` code from running, if the perpetrator managed to inject this into the HTML. Thought needs to be given to the most suitable & effective way to do this, be it via the Internet Explorer zone-based security controls in `Internet options`, using an allow-list of HTML tags akin to [htmlpurifier](http://htmlpurifier.org), or some other method.
* Currently named parameters are specified on the command line. A nice improvement might be to store these parameters in a settings XML file so that they do not need to be passed on every invocation of the script.
* Corrections should ideally be made so that the script falls into line with the [Powershell Style Guide](https://github.com/PoshCode/PowerShellPracticeAndStyle).

### Repository visibility

This GitHub repository is intentionally public, so that members of the Ex Libris Alma community may benefit from the script.
