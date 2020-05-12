# Dramatic.PSeDOCS
A PowerShell module to work with OpenText eDOCS DM.



# Introduction


# CmdLets

# Examples

```PowerShell


```

# Development and automated testing

## Development setup
The eDOCS installation will likely be 32-bit (without the 64-bit API installed). You'll have to use 32-bit PowerShell to develop and work with the Dramatic.PSeDOCS module and eDOCS.
During development, use TWO PowerShell ISE (x86) instances: one to develop the CmdLets and the other to do tests on the module. The testing ISE will remove and import the Dramatic.PSeDOCS module for every test, making sure the PowerShell runtime is clean of rogue variables and functions.

The Dramatic.PSeDOCS has been developed on a development Virtual Machine with eDOCS 5.3.1 (one library) and PowerShell 4.

## Automated testing
Pester is employed to do automated testing. Unit testing of the CmdLets is hard because the eDOCS COM components are very hard to mock.
But there are integration test scripts that test the CmdLets with one or more eDOCS environments.

One of the environments is the IVHO development VM with eDOCS 5.3.1. The integration test script is "tests\PSeDOCS.IVHO-VM.Integration.Tests.ps1". It tests the CmdLets with specific items from the eDOCS library.






# Team
- Coding: Victor Vogelpoel



