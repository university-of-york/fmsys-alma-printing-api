# Alma Printing With Powershell

#### Preamble
In recent years, Ex Libris — the Library systems supplier of the [Alma Library Management System](https://exlibrisgroup.com/products/alma-library-services-platform/) — have made significant improvements to the Alma printing stack. Alma is a web-based, cloud-hosted solution, that until recently relied upon customers' email for printing. Print jobs would be sent by email to a nominated customer Inbox representing the print queue. It would then be up to the customer to process and print the incoming emails, and while this works, this method of printing required some effort on the part of customers to implement a solution, and such solutions (particularly in our case) were fragile and prone to failure.

A few years back Ex Libris introduced two new methods of printing Alma printouts in addition to email. So-called "Quick printing", where the Alma web app passes the print data to the browser, and the browser receives a signal to invoke the browser print dialogue. This is useful for occasional ad hoc printing, and is assumed to have been made possible because of the availability of new printing APIs in modern web browsers. However, this is not useful for printing scenarios where a workflow is involved, because multiple clicks would be required, which would soon add up to an increase in time spent by the operator processing the items. This is where the second improvement made in recent years comes to the fore. The HTML print content can now be fetched via a new Alma Printing API, and [Ex Libris's Alma Print Daemon](https://github.com/ExLibrisGroup/alma-print-daemon) provides a client solution for customers to leverage this improvement. However, the Alma Print Daemon appears to have some downsides:

- being an Electron app it is very large; the installer is over 100MB!
- it is built on an old version of `node.js`, and I have tried & failed to repackage with `node.js` 16, due to a problem with `node-native-printer`
- it only supports standard paper sizes. We want to be able to print to `Roll Paper 80 x 297 mm` on an Epson TM series POS printer, and attempts to do this with the daemon results in very small printouts.

This repo contains my attempt to address the above downsides, although in its current state it has a few downsides of its own, which are explained later on.

#### What the script does

The Powershell script contained in this repo polls the Alma print queue for new jobs. If it finds any it will send the job to the specified printer.

#### Prerequisites
There are tasks that need to be done so that the Alma printer is available as an online queue, which are [explained here](https://knowledge.exlibrisgroup.com/Alma/Product_Documentation/010Alma_Online_Help_(English)/030Fulfillment/080Configuring_Fulfillment/020Fulfillment_Infrastructure/Configuring_Printers). For the purposes of running this script you will need to [get an API key](https://developers.exlibrisgroup.com/blog/how-to-set-up-and-use-the-alma-print-daemon/) to use.

#### How do I use the script?

Once you have an API key, you'll want to make it available to the script to use:
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

#### Future improvements

* Currently, if the script is interrupted while it is `Working..`, say by pressing `CTRL+C`, there's a chance that the original default printer and `Page Setup` settings as mentioned previously won't be restored. It might be possible to improve this by using `Try`,`Catch`,`Finally` [as indicated here](https://stackoverflow.com/a/15788979/1754517).
* The limitations of using Internet Explorer for printing could be overcome by using a third-party HTML rendering/printing tool [like this one](https://github.com/kendallb/PrintHtml). But sadly, having tested it, it doesn't cope with `Roll Paper 80 x 297 mm` paper size. Another option would be to pay for [Bersoft HTMLPrint](https://www.bersoft.com/htmlprint/), which would very likely work. Moving away from IE is probably a good thing, as [the availability of its COM object is in some doubt](https://techcommunity.microsoft.com/t5/windows-it-pro-blog/internet-explorer-11-desktop-app-retirement-faq/ba-p/2366549), following the annoucement of IE11's retirement, 15 June 2022.
* It would also be good to see if this script could be made into a Windows service, perhaps using `srvany` or [NSSM](https://nssm.cc/). Currently some `.cmd` files are provided as launchers. Such files can be put in the `shell:startup` directory so that the print queue checking is invoked immediately upon logon.
* Document all parameters in this README!
* It might be possible to avoid saving the HTML content to a file on disk by using [a technique like this one](https://stackoverflow.com/a/30642231). However, it's probably a useful contingency having the HTML persisted to disk in case re-prints are required, or — as mentioned in the previous section — if there is a problem printing the document.
