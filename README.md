# JPEGAutoRotator
A Windows context menu extension and command line tool for normalizing JPEG images that have EXIF orientation properties.  Images will be rotated accordingly and the EXIF orientation property removed.

I was motivated to create this extension after putting slideshows together for various software/hardware and discovering that the orientation property of JPEGs are not being respected.  This results in images being displayed upside down or sideways.  There are some tools that do this already, but not as seamlessly as a native context menu extension for Windows.

JPEGAutoRotator will appear in the context menu when you right-click single or multiple images (as well as folders that are treated as image folders).  The contents of sub folders are also enumerated and rotated as necessary.  JPEGAutoRotator will show progress UI if the operation takes more than a couple of seconds.  

The code takes advantage of systems with multiple processors/cores by parallelizing the work across multiple threads resulting in faster execution.

Rotation are lossess unless JPEGs do not have dimensions evenly divisible by 8.  Almost all modern digital cameras produce images of such size.
