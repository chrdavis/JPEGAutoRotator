# JPEGAutoRotator
A Windows context menu extension and command line tool for normalizing JPEG images that have EXIF orientation properties.  Images will be rotated accordingly and the EXIF orientation property removed.

I was motivated to create this extension after putting slideshows together for various software/hardware and discovering that the orientation property of JPEGs are not being respected.  This results in images being displayed upside down or sideways.  There are some tools that do this already, but not as seamlessly as a native context menu extension for Windows.

If you find this useful, feel free to buy me a beer!

[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donate_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=chrisdavis%40outlook%2ecom&lc=US&item_name=Chris%20Davis&item_number=JPEGAutoRotator&no_note=0&currency_code=USD&bn=PP%2dDonationsBF%3abtn_donate_LG%2egif%3aNonHostedGuest)

### Windows Shell Context Menu Extension
JPEGAutoRotator will appear in the context menu when you right-click single or multiple images (as well as folders that are treated as image folders).  The contents of sub folders are also enumerated and rotated as necessary.  JPEGAutoRotator will show progress UI if the operation takes more than a couple of seconds.  

The code takes advantage of systems with multiple processors/cores by parallelizing the work across multiple threads resulting in faster execution.

Rotation are lossess unless JPEGs do not have dimensions evenly divisible by 8.  Almost all modern digital cameras produce images of such size.

#### Screenshots

![Image description](/Images/JPEGAutoRotator_ContextMenu1.png)

![Image description](/Images/JPEGAutoRotator_ContextMenu_Progress.png)

### Command Line
A command line version (JPEGAutoRotator.exe) is also available which provides much more control over the operation that is performed.
```
USAGE:
  JPEGAutoRotator [Options] Path

Options:
  -?                    Displays this help message
  -Silent               No output to the console
  -Verbose              Detailed output of operation and results to the console
  -NoSubFolders         Do not enumerate items in sub folders of path
  -LosslessOnly         Do not rotate images which would lose data as a result.
  -ThreadCount          Number of threads used for the rotations. Default is the number of cores on the device.
  -Stats                Print detailed information of the operation that was performed to the console.
  -Progress             Print progress of operation to the console.
  -Preview              Provides a printed preview in csv format of what will be performed without modifying any images.

  Path                  Full path to a file or folder to auto-rotate.  Wildcards allowed.

Examples:
  > JPEGAutoRotator c:\users\chris\pictures\mypicturetorotate.jpg
  > JPEGAutoRotator c:\foo\
  > JPEGAutoRotator c:\bar\*.jpg
  > JPEGAutoRotator -ShowProgress -Stats c:\test\
  > JPEGAutoRotator -Preview -LosslessOnly -NoSubFolders c:\test\
```
