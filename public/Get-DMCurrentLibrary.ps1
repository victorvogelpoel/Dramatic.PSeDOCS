﻿# Get-DMCurrentLibrary.ps1
# 
# Aug 2017
# Copyright 2017 Dramatic Development
# If this works, it was written by Victor Vogelpoel (victor@victorvogelpoel.nl).
# If it doesn't, I don't know who wrote it.


function Get-DMCurrentLibrary
{
    [CmdletBinding()]
    param
    ()

    return $script:CurrentLibrary
}
