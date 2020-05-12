# Dramatic.PSeDOCS

A PowerShell module to work with OpenText eDOCS DM.
The PowerShell module uses the eDOCS (server) COM API to interact with eDOCS.

**THIS REPOSITORY IS ARCHIVED. IT IMPLEMENTS A VERY FEW POWERSHELL COMMANDS TO WORK WITH EDOCS. I DECIDED TO OPEN-SOURCE IT TO SERVE AS A LEARNING TOOL FOR DOING EDOCS STUFF WITH POWERSHELL. I had fun making it.**

**THIS POWERSHELL MODULE IS FAR FROM COMPLETE AND THE CODE IS FAR FROM PRODUCTION QUALITY. I WILL NOT BE DOING ANY MORE DEVELOPMENT ON IT. I WILL NOT ANSWER ANY QUESTIONS OR REQUESTS ON THIS SUBJECT. YOU'RE ON YOUR OWN.**


**THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.**

## Development setup

The eDOCS installation will likely be 32-bit (without the 64-bit API installed). You'll have to use 32-bit PowerShell to develop and work with the Dramatic.PSeDOCS module and eDOCS.

TIP: During development, use TWO PowerShell ISE (x86) instances: one to develop the CmdLets and the other to do tests on the module. The testing ISE will remove and import the Dramatic.PSeDOCS module for every test, making sure the PowerShell runtime is clean of rogue variables and functions.

The Dramatic.PSeDOCS has been developed on a development Virtual Machine with eDOCS 5.3.1 (one library) and PowerShell 4.

## Automated testing

Pester is employed to do automated testing. Unit testing of the CmdLets is hard because the eDOCS COM components are very hard to mock.
But there are integration test scripts that test the CmdLets with one or more eDOCS environments.

## Team

- Coding: Victor Vogelpoel
