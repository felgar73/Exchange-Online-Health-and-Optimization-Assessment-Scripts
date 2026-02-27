# HybridConfigCollection-Scripts
Used as part of Hybrid Assessment 

**Disclaimer**
This script is NOT an official Microsoft tool. Therefore use of the tool is covered exclusively by the license associated with this github repository.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. 
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

.Description
    Each script gathers a separate portion of data related to Exchange Hybrid including: mailflow connectors, free/busy sharing & OAuth configs.

.Notes
    Created by: Felix E. Garcia - felgar@microsoft.com

    Requirements:
    -Powershell should to be 'Run As Administrator'.

    General Notes:
    -Script assumes Kerberos Auth is enabled on-prem for remote Exchange session.
    -Requires EXO V3 Powershell module for EXO collection.
    -Requires Entra ID rights for SPN data.
    -On-Prem script should be run directly from and Exchange server via Exchange Mgmt Shell in order to get the best results.
    -If using remote powershell session, details collected for Exchange certificates will be limited due to a powershell limitation. 
    -Cloud collection script can be run from any machine that has the EXO powershell and MgGraph modules installed.

    -The script will first prompt you on whether you wish to collect on-premises Exchange data and whether or not a remote powershell session is needed for this (if running from local Exchange server, simply answer 'No' to this question). 
    -Cloud collection script will ask whether you wish to collect Exchange Online data (may remove this later). 
    -Once complete, it will display the location of collection data. Some files will be in '.xml' format and some in text files. In some cases, data is sent to both formats for flexibility.

    -The script will attempt to import the MgGraph powershell module in order to collect OAuth configs from Entra ID; if this fails it will continue with other operations but this data will require manual collection.
