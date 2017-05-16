Short Manual for GhostBuster

Ghostbuster is a program that lets you filter devices/device classes of which you want to remove the ghosted once.

For the advanced features of the latest beta (commandline and task scheduling) see Advanced Options.

Ghostbuster does exactly the same when you right click a device in the Windows Device Manager and choose uninstall. The ony difference is that GhostBuster does it in bulk for all filtered devices that are ghosted and thus saves a lot of time.

It will in contrast to the device manager GhostBuster will NOT uninstall active devices and certain device types that are considered to be services, even if they're ghosted.

Ghostbuster's filters are applied by right-clicking the devide list and select one of the 2 menu-items. Depending on the item chosen, parts of the device list will be colored (green not ghosted and red for ghosted devices).

GhostBuster_80445.jpg

In above screenshot the 'Generic USB Hub' is marked as a device to be removed if ghosted.
'USB Mass Storage Devices' are marked by using a wildcard ('* Mass Storage Devices'),
one of them is ghosted and will be removed.

When the Remove Ghosts button is pressed, the program might restart in elevated mode if needed by the operating system (Vista/Windows7).

If UAC is not present or the program is already running in elevated mode, Ghostbuster will remove ONLY the ghosted devices from the filtered (faint yellow colored) devices. It will NOT remove enties marked as services or non ghosted devices. The need for elevation is shown by a small UAC shield on the Remove Ghosts button.

Please be carefull with marking devices escpecially under System, Non PnP and Audio are some that seem to be ghosted permanently but are required for correct operation of hardware. Also do not remove (if present) the top entry with a null guid as name as I'm still not sure what it's purpose is. Use the class filter carefully.

Ghostbuster removes devices by name, class or wildcard so cannot be used to remove only one of two ghosted devices that share the same name, it will always try to remove all matching ghosted devices.

Ghostbuster will NOT remove drives or inf files so most drivers (like for flash drives) will re-install when the hardware is re-attached to the computer.
