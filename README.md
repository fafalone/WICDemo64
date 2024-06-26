# WICDemo64
Windows Imaging Component Demo


![image](https://github.com/fafalone/WICDemo64/assets/7834493/63732210-15fe-4949-b516-f7194f050f87)

Thought I'd see if my WIC demo worked in tB. Happy to report the initial import of the VB6 project worked flawlessly. From there I rewrote for 64bit. Thought it's worth noting how much easier WinDevLib API modules make this: Rather than go through and update all the Win API defs, all I did was comment them out, letting WinDevLib take over those in addition to taking over the COM interfaces from oleexp. After a minor hiccup (major bug in WinDevLib.IStream, now fixed), all that needed to be modified was removing some oleexp. qualifiers, switching `GetObject` to `GetObjectW` since the intrinsic version was getting priority, and changing handles/pointers to `LongPtr`.

Attached is the 64bit port, and a zip with the original VB6 project and unmodified tB import.


**Update (Version 0.4, 06 Apr 2024):** - Bug fixes to work with latest WDL versions ([Issue #1](https://github.com/fafalone/WICDemo64/issues/1)). Fix save notification position.

**Update (Version 0.3, 03 Mar 2024):** Added BMP save option; also .twinproj has been updated to use a more recent version of WinDevLib (formerly tbShellLib) due to errors in the package tB did not raise at the time this project was released.

Original project page:
[[VB6] Intro to the Windows Imaging Component (WIC): Scale and convert JPG to PNG](https://www.vbforums.com/showthread.php?879907) (the attached version adds BMP as a supported save format)
