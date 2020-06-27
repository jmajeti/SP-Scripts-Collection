/*
This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that You agree:
(i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and 
(iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
*/

USE WSS_Content_MissingAssembly_Snapshot

SELECT WebID, SiteID, DirName, LeafName, tp_WebPartTypeId, tp_WebPartIDProperty
FROM AllDocs WITH (NOLOCK) inner join AllWebParts on AllDocs.Id = AllWebParts.tp_PageUrlID
WHERE AllWebParts.tp_WebPartTypeID = '817a5bd9-0698-4c62-7c51-5c9966468f0d'

