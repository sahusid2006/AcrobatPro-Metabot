/*ADOBE SYSTEMS INCORPORATED
 Copyright (C) 1994-2006 Adobe Systems Incorporated
All rights reserved.

 NOTICE: Adobe permits you to use, modify, and distribute this file
 in accordance with the terms of the Adobe license agreement
 accompanying it.
------------------------------------------------------------

BasicIacCS
- This is a simple Acrobat IAC C# code to perform certain functions.
'------------------------------------------------------------*/
using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Collections.Generic;
using System.Data;
using Acrobat;

namespace AutomationAnywhere
{
    /// <summary>
    /// Summary description for BasicIac.
    /// </summary>
    public class BasicIac
    {

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        void Dispose(bool disposing)
        {
            if (disposing)
            {

            }

        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        

        //Function to open the PDF and get the number of pages
        public int StartAcrobatIac(String szPdfPathConst)
        {
            //variables
            int iNum = 0;

            try
            {
                //IAC objects
                CAcroPDDoc pdDoc;
                CAcroAVDoc avDoc;
                CAcroApp avApp;


                //set AVApp Project
                avApp = new AcroAppClass();

                //set AVDoc object
                avDoc = new AcroAVDocClass();

                //open the PDF
                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //set the pdDoc object and get some data
                    pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
                    iNum = pdDoc.GetNumPages();
                }
                else
                {
                    iNum = 0;
                }
            }
            catch (Exception)
            {
                iNum = 0;
            }
            return iNum;
        }

        //Function to Check if Word is present or not.
        public bool IsWordPresent(string szPdfPathConst, string searchword)
        {
            //variables
            bool TextCheck;

            try
            {
                //IAC objects
                CAcroPDDoc pdDoc;
                CAcroAVDoc avDoc;
                CAcroApp avApp;

                //set AVApp Project
                avApp = new AcroAppClass();

                //set AVDoc object
                avDoc = new AcroAVDocClass();

                //Show Acrobat
                avApp.Show();

                //open the PDF if it isn't already open

                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //set the pdDoc object and get some data
                    pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
                    //Checking if word is present or not
                    TextCheck = avDoc.FindText(searchword, 0, 0, 0);
                }
                else
                {
                    TextCheck = false;
                }
            }

            catch (Exception)
            {
                TextCheck = false;
            }
            return TextCheck;
        }


        //////////////////////////Function to Find Text and Obtain Corresponding Pages///////////////////
        public string GetPageNumforWord(string szPdfPathConst, string searchword, int bCaseSensitive, int bWholeWordsOnly)
        {
            //Initializing variables
            int iNum = 0;
            bool TextCheck;
            int PageNum;
            bool GoToStatus;
            string PageNumConsol = "";
            int ScanPage;
            List<int> PageList = new List<int>();
            List<string> PageListString = new List<string>();

            try
            {
                //Declaring relevant IAC objects
                CAcroPDDoc pdDoc;
                CAcroAVDoc avDoc;
                CAcroApp avApp;
                CAcroAVPageView avPage;

                //set AVApp Project
                avApp = new AcroAppClass();

                //set AVDoc object
                avDoc = new AcroAVDocClass();

                //open the PDF if it isn't already open

                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //set the pdDoc object and get some data
                    pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();

                    //Getting Total Number of Pages in the PDF
                    iNum = pdDoc.GetNumPages();

                    //set AVPage View object
                    avPage = (CAcroAVPageView)avDoc.GetAVPageView();

                    //Navigating to Page 1 to initiate search
                    ScanPage = 0;
                    GoToStatus = avPage.GoTo(ScanPage);

                    //Checking if word is present or not
                    TextCheck = avDoc.FindText(searchword, bCaseSensitive, bWholeWordsOnly, 0);

                    //Declaring variable for storing the previous page number 
                    int PageNumPrev = 0;

                    if (TextCheck == true)
                    {
                        PageNum = avPage.GetPageNum();
                        //First Page is 0 and thus offset is being taken care of
                        PageNum = PageNum + 1;
                        PageList.Add(PageNum);

                        //Incrementing Page numbers and searching for more instances
                        while (TextCheck == true)
                        {
                            //Going to the page next to the previous search result - Not incremented by 1 since PageNum was already incremented for recording.
                            ScanPage = PageNum;
                            if (ScanPage == iNum)
                            {
                                TextCheck = false;
                                break;
                            }
                            GoToStatus = avPage.GoTo(ScanPage);
                            TextCheck = avDoc.FindText(searchword, bCaseSensitive, bWholeWordsOnly, 0);
                            PageNum = avPage.GetPageNum();

                            //Exit loop in case the previous page number is bigger than the current
                            if (PageNumPrev > PageNum)
                            {
                                break;
                            }
                            //Assigning the page number for this search iteration to a previous variable
                            PageNumPrev = PageNum;

                            //First Page is 0 and thus offset is being taken care of
                            PageNum = PageNum + 1;
                            PageList.Add(PageNum);

                        }
                    }
                    else
                    {
                        PageNum = 0;
                        PageList.Add(PageNum);
                    }
                }
                else
                {
                    PageNum = 0;
                    PageList.Add(PageNum);
                }

                //Removing Duplicates in the list due to multiple occurences of word on the same page
                List<int> PageListFilter = new List<int>();
                foreach (int i in PageList)
                {
                    if (!PageListFilter.Contains(i))
                    {
                        PageListFilter.Add(i);
                    }
                }

                //Converting Integer List for Page List to String List
                PageListString = PageListFilter.ConvertAll<string>(delegate (int i)
                {
                    return i.ToString();
                });

                //Converting String List to Comma Delimited List
                PageNumConsol = string.Join(",", PageListString.ToArray());
            }
            catch(Exception)
            {
                PageNumConsol = "Unknown Exception";
            }

            return PageNumConsol;
        }

        ////////////////////////////////SAVING PDF/////////////////////////////////////////////
        public bool SavePDF(string szPdfPathConst, string sFullPath)
        {
            //Declaring Variables
            bool SaveAs;

            try
            {
                //IAC objects
                CAcroPDDoc pdDoc;
                CAcroAVDoc avDoc;

                //set AVDoc object
                avDoc = new AcroAVDocClass();

                //open the PDF
                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //set the pdDoc object and get some data
                    pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
                    SaveAs = pdDoc.Save(1, sFullPath);
                }
                else
                {
                    SaveAs = false;
                }
            }
            catch
            {
                SaveAs = false;
            }

            //Returning output var
            return SaveAs;

        }

        /// <summary>
        /// //////////////////////CLOSING PDF/////////////////////////////////////////
        /// </summary>

        public bool ClosePDFNoChanges(string szPdfPathConst)
        {
            //Initializing Variables
            bool CloseCheck;

            try
            {
                //SettingObject
                CAcroApp avApp;
                CAcroAVDoc avDoc;

                //set AVApp Project
                avApp = new AcroAppClass();
                //set AVDoc object
                avDoc = new AcroAVDocClass();

                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //Checking if word is present or not
                    CloseCheck = avDoc.Close(1);
                }
                else
                {
                    CloseCheck = false;
                }

                avApp.CloseAllDocs();
                avApp.Exit();
            }
            catch
            {
                CloseCheck = false;
            }

            return CloseCheck;

        }
    }
}
