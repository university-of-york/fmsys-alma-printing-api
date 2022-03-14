# Alma Printing With Powershell

#### Preamble
In recent years, Ex Libris — the Library systems supplier of the [Alma Library Management System](https://exlibrisgroup.com/products/alma-library-services-platform/) — have made significant improvements to the Alma printing stack. Alma is a web-based, cloud-hosted solution, that until recently relied upon customers' email for printing. Print jobs would be sent by email to a nominated customer Inbox representing the print queue. It would then be up to the customer to process and print the incoming emails, and while this works, this method of printing required some effort on the part of customers to implement a solution, and such solutions (particularly in our case) were fragile and prone to failure.

A few years back Ex Libris introduced two new methods of printing Alma printouts in addition to email. So-called "Quick printing", where the Alma web app passes the print data to the browser, and the browser receives a signal to invoke the browser print dialogue. This is useful for occasional ad hoc printing, and is assumed to have been made possible because of the availability of new printing APIs in modern web browsers. However, this is not useful for printing scenarios where a workflow is involved, because multiple clicks would be required, which would soon add up to an increase in time spent by the operator processing the items. This is where the second improvement made in recent years comes to the fore. The HTML print content can now be fetched via a new Alma Printing API, and [Ex Libris's Alma Print Daemon](https://github.com/ExLibrisGroup/alma-print-daemon) provides a client solution for customers to leverage this improvement. However, the Alma Print Daemon appears to have some downsides:

- being an Electron app it is very large; the installer is over 100MB!
- it is built on an old version of `node.js`, and I have tried & failed to repackage with `node.js` 16, due to a problem with the dependency `node-native-printer`
- it only supports standard paper sizes. We want to be able to print to `Roll Paper 80 x 297 mm` on an Epson TM series POS printer, and attempts to do this with the daemon results in very small printouts

This repo contains my attempt to address the above downsides, although in its current state it has a few downsides of its own, which are explained later on.

#### What the script does

The Powershell script contained in this repo polls the Alma print queue for new jobs. If it finds any it will send the job to the specified printer and mark the print job as `Printed`.

#### Prerequisites
There are tasks that need to be done in the Alma web UI Configuration area, so that the Alma printer is available as an online queue, which are [explained here](https://knowledge.exlibrisgroup.com/Alma/Product_Documentation/010Alma_Online_Help_(English)/030Fulfillment/080Configuring_Fulfillment/020Fulfillment_Infrastructure/Configuring_Printers). For the purposes of running this script you will need to [get an API key](https://developers.exlibrisgroup.com/blog/how-to-set-up-and-use-the-alma-print-daemon/) to use.

#### How do I use the script?

Once you have an API key, you'll want to make it available for the script to use:
```
Set-Location <script-dir>
. .\FetchAlmaPrint.ps1;Invoke-Setup
```
You will be prompted to enter your API key, which you can paste in. This is then stored in an XML file for future use by the script.

The second thing you'll want to do is determine the Alma print queue to monitor:

```
 . .\FetchAlmaPrint.ps1;Fetch-Printers
 ```
 This will list all of the Alma printers. You'll want to note the Printer ID for the queue you're interested in printing from.

 (note that you only need to dot source the script if it's no longer in memory, i.e. you closed Powershell between invocations)

 Thirdly, and lastly, you'll want to run the `Fetch-Jobs` function, with the parameters set according to the requirements, e.g.:
```
 . .\FetchAlmaPrint.ps1;Fetch-Jobs -printerId "848838010001381" -localPrinterName "EPSON TM-T88III Receipt" -checkInterval 20
 ```
#### How does the script work?

The script works by running a loop to check the queue every `X` seconds. The number of seconds can be changed with the `-checkInterval` parameter of the `Fetch-Jobs` function. By default, this is `30` seconds. the communications with Alma's API is achieved with Powershell's `Invoke-RestMethod`.

The HTML content is saved to a file stored in the `tmp_printouts` subdirectory. Each file contains the content for one printout, and the filename will contain the `letterId` (the ID assigned by Alma to the printout). If something related to printing goes wrong, but the "fetch" from Alma succeeded, it would then be possible to manually print the content from the stored file.

Rendering and printing the HTML is achieved using an Internet Explorer COM object. This is probably the weakest area of the script; there is currently no way to internally specify which printer to print to, so the script relies upon setting the printer specified by `-localPrinterName` as the Windows default printer while the script runs, restoring the default printer after each run. The script also does something quite similar for the `Page Setup` (margins, etc), since these are not printer-specific, but typically we'll want different values according to the printer we want to print to.

By default, the script will only fetch printouts that have the Alma printout status `Pending`. Once the printout has been printed, the script changes the status to `Printed`. There is a `Fetch-Jobs` function parameter `-printStatuses` that can be added, which can be used to fetch printouts with other statuses (`Printed`, `Pending`, `Canceled`, `ALL`). For example, `-printStatuses "Canceled"`

#### Deployment
In a production environment, it is recommended that two identical script shortcut (`.lnk`) files are created. These should be put in the `shell:startup` & `shell:desktop` directories. The shortcut in `shell:startup` ensures the script runs automatically upon logon, and the shortcut in `shell:desktop` provides an option to manually invoke the script, if the Powershell window was accidentally closed. The script will run minimised, so as to run discretely in the background, lessening the chance of an operator accidentally closing the window.

To create the shortcuts, take the following steps:

1. Assuming this repo is in `C:\fmsys-alma-printing-api`, create a new shortcut in this folder:
   1. In File explorer, right-click in the whitespace, and from the menu left-click `New` > `Shortcut`
   2. In the resulting wizard, enter `powershell` for the location of the item. Click the `Next` button
   3. In the next screen, give the shortcut a suitable name like `Alma Slip Printing`. Click the `Finish` button


2. Now make some modifications to your new shortcut file
   1. Right-click the shortcut file, and from the menu left-click `Properties`
   2. In the `Shortcut` tab of the resulting dialogue, edit the following fields:

Target: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoLogo -Command %ALMA_PRINTING_CMD%`  
Start in: `C:\fmsys-alma-printing-api`  
Run: `Minimized`  

3. Click the `OK` button to save these changes.
4. In the Windows `Run` app, enter `shell:common startup`. Explorer will open the startup folder.
5. Move your shortcut from `C:\fmsys-alma-printing-api` to this folder (_admin rights required_)
6. In the Windows `Run` app, enter `shell:common desktop`. Explorer will open the desktop folder
7. From `shell:common startup`, copy your shortcut to the `shell:common desktop` folder (_admin rights required_)
8. Add a new System environment variable (_admin rights required_):
   1. In the Windows `Run` app, enter `rundll32 sysdm.cpl,EditEnvironmentVariables` and `CTRL+SHIFT`click the `OK` button to run the aforementioned elevated
   2. In the lower pane add a new System environment variable by clicking the `New` button and then adding the following:
   3. Variable name: `ALMA_PRINTING_CMD`, Variable value:


 ```
 "& { Start-Sleep 30;. .\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId '<printerId>' -localPrinterName '<printerName>' }"
 ```

Lastly, reboot the PC to check that the script is starting normally.

###### Notes

- A script `DeployShortcuts.ps1` has been added to the `extras` subfolder, to make it easier to deploy to production environments. It scripts the above steps. Run as Administrator.
- If the script _flashes by_, i.e. errors appear, but the Powershell window closes too quickly to see what caused it, try temporarily adding the `-NoExit` Powershell option to the shortcut target.
- Note the `Start-Sleep 30;` bit. It was noted that if the script starts too quickly after logging in, then you might see errors. This ensures a delay before starting the queue checking.
- If you're wondering why step 8 is required i.e. why we don't embed the parameters directly in the shortcut, this is because there is a 260 character length limit on shortcuts' `Target` field, and this will very likely be exceeded by including additional parameters.

A list of UoY `ALMA_PRINTING_CMD` variable values is as follows:

**Interlending receiving**
```
"& { Start-Sleep 30;. .\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId '19195349880001381' -localPrinterName 'PUSH_ITSPRN0705 [Harry Fairhurst - Information Services LFA/ LFA023](Mobility)' -marginTop '0.3' -jpgBarcode }"
```

**JBM Holds processing**
```
"& { Start-Sleep 30;. .\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId '848838010001381' -localPrinterName 'EPSON TM-T88III Receipt' }"
```

**KML Holds processing**
```
"& { Start-Sleep 30;. .\FetchAlmaPrint.ps1;Fetch-Jobs -checkInterval 15 -printerId '993537480001381' -localPrinterName 'EPSON TM-T88III Receipt' }"
```
###### Barcode unreadable?
A problem was identified with the readability of the barcodes when printed using the script. The key points about this problem are:
- the printed barcodes appeared to be missing their right guard bar
- when the HTML content is rendered on screen in Internet Explorer, from the files stored in `tmp_printouts`, the guard bar is _not_ missing
- the Opticon branded scanners in use are incapable of reading the "faulty" printed barcodes, however, they can be read using ones' smartphone
- the problem is not present when the HTML content is printed from an alternative web browser such as Google Chrome
- the problem was narrowed down to it being Internet Explorer related, and there is [a very similar sounding problem described here](https://social.technet.microsoft.com/Forums/windows/en-US/9276a5b1-24cf-4973-873c-768068617e79/issue-printing-with-internet-explorer-10-11?forum=ieitprocurrentver), which pinpoints the XPS subsystem (that IE uses for printing) as the root cause
- after experimentation, it was found that the barcodes could be converted from `PNG` to `JPG` format in order to resolve the readability problem

To provide a solution for this problem, a new `base64Png2Jpg` function was added which converts the base64 PNG data to base64 JPG. This can be used by adding the `Fetch-Jobs` function switch parameter `-jpgBarcode`.

###### IE First-Run

Because the script relies upon Internet Explorer for rendering & printing the HTML, it is likely you'll see the following first-run box:

![An image of the Internet Explorer 11 first-run box](./extras/IE11_First_Run_Image.png?raw=true)

To prevent this box from reappearing, a helper script is provided. Perform the following one-time step in an elevated Powershell window:
```
Set-Location extras
.\DisableFirstRunIE.ps1
```

###### tmp_printouts housekeeping

The `tmp_printouts` directory will accumulate many HTML files over time. It's probably a useful contingency having the HTML files persisted to disk in case re-prints are required, or if there is a problem printing the document. However, over time these files may consume a significant amount of disk space, so it's recommended to delete the older ones while keeping the more recent ones. To help automate this housekeeping process, a Task Scheduler XML template is provided which leverages `forfiles` to delete files older than 30 days. This can be adjusted according to local needs; just edit the XML before importing, or modify the task once imported. See `extras/fmsys-alma-printing-api - clean tmp_printouts directory.xml`

#### Future improvements

* Document all parameters in this README!
* Currently, if the script is interrupted while it is `Working..`, say by pressing `CTRL+C`, there's a chance that the original default printer and `Page Setup` settings as mentioned previously won't be restored. It might be possible to improve this by using `Try`,`Catch`,`Finally` [as indicated here](https://stackoverflow.com/a/15788979/1754517).
* The limitations of using Internet Explorer for printing could be overcome by using a third-party HTML rendering/printing tool [like this one](https://github.com/kendallb/PrintHtml). But sadly, having tested it, it doesn't cope with `Roll Paper 80 x 297 mm` paper size. Another option would be to pay for [Bersoft HTMLPrint](https://www.bersoft.com/htmlprint/), which would very likely work. Moving away from IE is probably a good thing, as [the availability of its COM object is in some doubt](https://techcommunity.microsoft.com/t5/windows-it-pro-blog/internet-explorer-11-desktop-app-retirement-faq/ba-p/2366549), following the announcement of IE11's retirement, 15 June 2022.
* It would also be good to see if this script could be made into a Windows service, perhaps using `srvany` or [NSSM](https://nssm.cc/), instead of invoking the script via shortcuts.
* To protect against a possible API endpoint security compromise, it would be a good idea to sanitise the HTML letter content before "opening" it, as is effectively done with `$ie.Navigate($printOut)`. The idea would be that this would protect against e.g. malicious `<script></script>` code from running, if the perpetrator managed to inject this into the HTML. Thought needs to be given to the most suitable & effective way to do this, be it via the Internet Explorer zone-based security controls in `Internet options`, using an allow-list of HTML tags akin to [htmlpurifier](http://htmlpurifier.org), or some other method.
* Currently named parameters are specified on the command line. A nice improvement might be to store these parameters in a settings XML file so that they do not need to be passed on every invocation of the script. This would help reduce the number of steps to deploy the script in production environments.
