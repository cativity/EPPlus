/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/10/2023         EPPlus Software AB       Initial release EPPlus 6.2
 *************************************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.ExcelXMLWriter;

internal class ExtLstHelper
{
    List<string> listOfExts = new List<string>();
    Dictionary<string, int> uriToIndex = new Dictionary<string, int>();

    public ExtLstHelper(string xml) => this.ParseIntialXmlToList(xml);

    private void ParseIntialXmlToList(string xml)
    {
        int start = 0,
            end = 0;

        GetBlock.Pos(xml, "extLst", ref start, ref end);

        //If the node isn't just a placeholder
        if (end - start > 10)
        {
            int contentStart = start + "<ExtLst>".Length;
            string extNodesOnly = xml.Substring(contentStart, end - contentStart - "</ExtLst>".Length);

            string[] strLst = { "</ext>" };
            this.listOfExts = extNodesOnly.Split(strLst, StringSplitOptions.RemoveEmptyEntries).ToList();

            for (int i = 0; i < this.listOfExts.Count; i++)
            {
                int startOfUri = this.listOfExts[i].LastIndexOf("{");
                int endOfUri = this.listOfExts[i].LastIndexOf("}") + 1;

                string uri = this.listOfExts[i].Substring(startOfUri, endOfUri - startOfUri);

                this.uriToIndex.Add(uri, i);
                this.listOfExts[i] += "</ext>";
            }
        }
    }

    /// <summary>
    /// Inserts content after the uriNode
    /// Note that this is only intended to be done once per type of node and it will throw error
    /// if the same uri is attempted in two separate calls or if it's already been read in initally.
    /// </summary>
    /// <param name="uri"></param>
    /// <param name="content"></param>
    /// If <param name="uriOfNodeBefore"> is blank sets content as the first ext</param>
    internal void InsertExt(string uri, string content, string uriOfNodeBefore)
    {
        int indexOfNode = -1;

        if (uriOfNodeBefore != "")
        {
            indexOfNode = this.uriToIndex[uriOfNodeBefore];
        }

        List<string> keys = new List<string>(this.uriToIndex.Keys);

        if (indexOfNode == -1)
        {
            this.listOfExts.Insert(0, content);

            foreach (string key in keys)
            {
                this.uriToIndex[key] += 1;
            }

            this.uriToIndex.Add(uri, 0);
        }
        else
        {
            if (indexOfNode + 1 > this.listOfExts.Count)
            {
                this.listOfExts.Add(content);
            }
            else
            {
                this.listOfExts.Insert(indexOfNode + 1, content);

                foreach (string key in keys)
                {
                    if (indexOfNode + 1 >= this.uriToIndex[key])
                    {
                        this.uriToIndex[key] += 1;
                    }
                }
            }

            this.uriToIndex.Add(uri, indexOfNode + 1);
        }
    }

    internal string GetWholeExtLst()
    {
        string extLstString = "<extLst>";

        for (int i = 0; i < this.listOfExts.Count; i++)
        {
            extLstString += this.listOfExts[i];
        }

        extLstString += "</extLst>";

        return extLstString;
    }
}