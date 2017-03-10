using System;
using System.Web;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using System.ComponentModel;
using System.Xml;
using System.Configuration;
using System.IO;
using System.Drawing;
using System.Collections.Generic;
using System.Diagnostics;
using Adaptive;
using Adaptive.Data;
using Adaptive.Service;
using Adaptive.Archive;
using C1.C1Excel;
using Microsoft.Win32;
using C1.Win.C1FlexGrid;
using Adaptive.Creation;
using Adaptive.Lifecycle;
using Adaptive.Windows.View;
using Adaptive.View;
using Adaptive.Category;
using Adaptive.Relationship;
using Adaptive.Windows.Action;
using Adaptive.Action;


namespace Custom_Export
{
    public class Custom_Export
    {

        public Custom_Export()
        {
        }
        public void ECOFlow(string entityId, string userid)
        {
             BottomUpRevisionForm ps = new BottomUpRevisionForm();



            ps.PartID = entityId;
            ps.ShowDialog();

            // 부품 개정시 관계된 상위 어셈블리를 리스트화 
            string xmldocs = string.Empty;
            string xmldocsall = string.Empty;

            if (ps.DialogResult == DialogResult.OK)
            {
                foreach (ObjectInfo objInfo in ps.SelectedParents)
                {
                    //MessageBox.Show("s1");
                    DataTable topsPartName =  Services.ApplicationServices.StructureSvc.GetTops(entityId, "EBOM", "Part");
                    //MessageBox.Show("s2");
                    
                    foreach (DataRow topPartName in topsPartName.Rows)
                    {
                        //MessageBox.Show(Convert.ToString(objInfo.ObjectID) + " : " + Convert.ToString(objInfo.Version) + " : " + Convert.ToString(objInfo.Creator));
                        xmldocs = "\r\n Name : " + Convert.ToString(topPartName["P$Name"]) + ", Version : " + Convert.ToString(topPartName["P$Version"]) + ", 작성자 : ";
                        xmldocsall = xmldocsall + xmldocs;
                    }
                    MessageBox.Show(xmldocsall);
                }
                MessageBox.Show("성공");
            }
            //Services.ApplicationServices.CreateSvc

        }
        public void PmStructure(string entityId, string userId)
        {

            //MessageBox.Show("1");
            int i = entityId.IndexOf('.');
            string PartClassname = entityId.Substring(0, i);
            int totalSeq = 1;
            int baseNumbering = 0;
            string pUrid = string.Empty;
            //MessageBox.Show("1");
            string newNumber = Services.ApplicationServices.DataSvc.ExecuteScalar("SELECT Max(serial) FROM numbering WHERE master_type='Part' AND object_type='PartPM' AND category_id='P$PARTPM-1075' AND used=1");
            if (newNumber == "")
            {
                baseNumbering = 1;
            }
            else
            {
                baseNumbering = Convert.ToInt16(newNumber) + 1;
            }
            Services.ApplicationServices.DataSvc.ExecuteScalar("INSERT INTO numbering (urid, master_type, object_type, category_id, serial, used, reserved) VALUES ('-1', 'Part', 'PartPM', 'P$PARTPM-1075', " + baseNumbering + ", 1, 1)");

            int newNumbering = 1000000 + baseNumbering;
            newNumber = Convert.ToString(newNumbering).Substring(1, 6);
            string Totalentityid = Services.ApplicationServices.DataSvc.ExecuteScalar("SELECT urid FROM part_info where urid like '" + PartClassname + "%' and P$Name  = 'total' and latest = '1'");

            string ModelName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$ModelName");
            CreationBuilder cb = new CreationBuilder("Part", "PartPM", "P$PARTPM-1075", userId);
            cb.Append("P$Number", "A01-" + newNumber);
            cb.Append("P$Name", "SET ASS'Y");
            cb.Append("P$ModelName", ModelName);

            Services.ApplicationServices.CreateSvc.CreateObject(cb.ToXml());
            string partEntityId = Services.ApplicationServices.DataSvc.ExecuteScalar("select*from part_info where P$number = 'A01-" + newNumber + "' and P$name like 'SET_ASS%'");


            Services.ApplicationServices.DataSvc.ExecuteScalar(" UPDATE numbering SET reserved=0, urid='" + partEntityId + "' WHERE master_type='Part' AND object_type='PartPM' AND category_id='P$PARTPM-1075' AND serial = " + newNumber);
            //MessageBox.Show("2");
            Services.ApplicationServices.StructureSvc.ClearStructure(entityId);
            
            if(Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, partEntityId) != true) //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
            {
            totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);  
            string setStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userId, totalSeq, partEntityId, 1);
            }

            //MessageBox.Show(PartClassname);ruytjghjg
       

            //string Totalentityid = Services.ApplicationServices.DataSvc.ExecuteScalar("SELECT urid FROM part_info where urid like '" + PartClassname + "%' and P$Name  = 'total' and latest = '1'");
            DataTable rowSetDt = Services.ApplicationServices.StructureSvc.GetTopDownStructure(Totalentityid, "EBOM","Part");
            //MessageBox.Show("3");
           

            foreach (DataRow datarow2 in rowSetDt.Rows)
            {

                string pName = Convert.ToString(datarow2["P$Name"]);
                if ("SET ASS'Y" == Convert.ToString(datarow2["P$Name"]) && "1" == Convert.ToString(datarow2["tree_level"]))
                {
                    //MessageBox.Show("4");
                    pUrid = Convert.ToString(datarow2["child_id"]);
                    PMLogic(pUrid, partEntityId, userId, entityId);

                }
            }

            if (pUrid == string.Empty || Totalentityid == string.Empty)
            {
                if (Totalentityid == string.Empty)
                {
                    MessageBox.Show("TOTAL 제품이 없습니다.");
                }
                else
                {
                    if (pUrid == string.Empty)
                    {
                        MessageBox.Show("SET ASS'Y 어셈블리가 없습니다.");
                    }
                }
            }
            else
            {
                //MessageBox.Show("5");
                PMLogic(Totalentityid, entityId, userId, entityId);

                MessageBox.Show("BOM 자동화 구성 완료.");
            }
           
        }


        public void PMLogic(string pUrid, string partEntityId, string userId, string entityId)
        {
            try
            {
                int totalSeq = 1;
                int maxOptionCount = 0;
                int matchCount = 0;
                string[] optSeOption = new string[10];
                string[] PartOpt = new string[10];
                string[] dataRowOption = new string[10];

                PartOpt[0] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Optone");//설계 변경 통보의 번호
                PartOpt[1] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Opttwo");//설계 변경 통보의 번호
                PartOpt[2] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Optthr");//설계 변경 통보의 번호
                PartOpt[3] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Optfour");//설계 변경 통보의 번호
                PartOpt[4] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Optfive");//설계 변경 통보의 번호
                PartOpt[5] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Optsix");//설계 변경 통보의 번호
                PartOpt[6] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Optsev");//설계 변경 통보의 번호
                PartOpt[7] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Opteig");//설계 변경 통보의 번호
                PartOpt[8] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Optnine");//설계 변경 통보의 번호
                PartOpt[9] = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Optten");//설계 변경 통보의 번호

                //MessageBox.Show("6" + pUrid + partEntityId + userId);


                DataTable rowDt = Services.ApplicationServices.StructureSvc.GetTopDownStructure(pUrid, "EBOM", "part");


                foreach (DataRow datarow in rowDt.Rows)
                {
                    if ("1" == Convert.ToString(datarow["tree_level"]))
                    {
                        maxOptionCount = 0;
                        matchCount = 0;
                        dataRowOption[0] = Convert.ToString(datarow["S$Optone"]);
                        dataRowOption[1] = Convert.ToString(datarow["S$Opttwo"]);
                        dataRowOption[2] = Convert.ToString(datarow["S$Optthr"]);
                        dataRowOption[3] = Convert.ToString(datarow["S$Optfour"]);
                        dataRowOption[4] = Convert.ToString(datarow["S$Optfive"]);
                        dataRowOption[5] = Convert.ToString(datarow["S$Optsix"]);
                        dataRowOption[6] = Convert.ToString(datarow["S$Optsev"]);
                        dataRowOption[7] = Convert.ToString(datarow["S$Opteig"]);
                        dataRowOption[8] = Convert.ToString(datarow["S$Optnine"]);
                        dataRowOption[9] = Convert.ToString(datarow["S$Optten"]);

                        //BOM 구조 하나의 옵션 개수
                        for (int optCount = 0; optCount < 10; optCount++)
                        {

                            if (dataRowOption[optCount] != string.Empty)
                            {
                                //MessageBox.Show(dataRowOption[optCount].Trim().ToLower());
                                maxOptionCount++;
                            }


                        }

                        //MessageBox.Show("구조 개수 끝" + Convert.ToString(maxOptionCount));
                        if (maxOptionCount != 0)
                        {
                            if (dataRowOption[0].ToLower() == "all")//BOM의 옵션중 ALL인경우에는 무조건 복사.
                            {


                                if (Services.ApplicationServices.StructureSvc.IsExistStructure(partEntityId, "EBOM", -1, Convert.ToString(datarow["child_id"])) != true)  //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                {
                                    //MessageBox.Show("구조 체크 나간 후");
                                    totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(partEntityId, 1, 1);
                                    string structureId = Services.ApplicationServices.StructureSvc.CreateStructure(partEntityId, "EBOM", userId, totalSeq, Convert.ToString(datarow["child_id"]), 1);
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optone", dataRowOption[0]);
                                    //MessageBox.Show(structureId);
                                }

                            }
                            else
                            {

                                for (int optPartCount = 1; optPartCount < 10; optPartCount++)
                                {
                                    int optIndex = dataRowOption[optPartCount].IndexOf(',');

                                    if (optIndex > 0)
                                    {
                                        optSeOption = dataRowOption[optPartCount].Split(',');

                                        for (int k = 0; k < optSeOption.Length; k++)
                                        {

                                            if (PartOpt[optPartCount].Trim().ToLower() == optSeOption[k].Trim().ToLower() && PartOpt[optPartCount].Trim() != string.Empty)
                                            {

                                                matchCount++;

                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (PartOpt[optPartCount].Trim().ToLower() == dataRowOption[optPartCount].Trim().ToLower() && PartOpt[optPartCount].Trim() != string.Empty)
                                        {

                                            matchCount++;

                                        }
                                    }

                                    //MessageBox.Show("구조 매칭 끝" + Convert.ToString(matchCount));
                                    optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                    //MessageBox.Show(Convert.ToString(datarow["P$Name"]) + " : " + PartOpt[optPartCount].Trim() + " : " + dataRowOption[optPartCount].Trim());
                                }

                                if (matchCount == maxOptionCount && maxOptionCount != 0)
                                {


                                    if (Services.ApplicationServices.StructureSvc.IsExistStructure(partEntityId, "EBOM", -1, Convert.ToString(datarow["child_id"])) != true)  //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                    {
                                        //MessageBox.Show("구조 체크 나간 후");
                                        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(partEntityId, 1, 1);
                                        string structureId = Services.ApplicationServices.StructureSvc.CreateStructure(partEntityId, "EBOM", userId, totalSeq, Convert.ToString(datarow["child_id"]), 1);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optone", dataRowOption[0]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Opttwo", dataRowOption[1]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optthr", dataRowOption[2]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optfour", dataRowOption[3]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optfive", dataRowOption[4]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optsix", dataRowOption[5]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optsev", dataRowOption[6]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Opteig", dataRowOption[7]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optnine", dataRowOption[8]);
                                        Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(structureId, "S$Optten", dataRowOption[9]);
                                        //MessageBox.Show(structureId);
                                    }


                                }




                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptOne.Trim() == optSeOption[k].Trim() && PartOptOne.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;



                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOpttwo.Trim() == optSeOption[k].Trim() && PartOpttwo.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;



                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptthr.Trim() == optSeOption[k].Trim() && PartOptthr.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;



                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptFour.Trim() == optSeOption[k].Trim() && PartOptFour.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;




                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptFive.Trim() == optSeOption[k].Trim() && PartOptFive.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;


                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptSix.Trim() == optSeOption[k].Trim() && PartOptSix.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;



                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptSev.Trim() == optSeOption[k].Trim() && PartOptSev.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;


                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptEig.Trim() == optSeOption[k].Trim() && PartOptEig.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;

                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptNine.Trim() == optSeOption[k].Trim() && PartOptNine.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;


                                //optSeOption = dataRowOption.Split(',');

                                //for (int k = 0; k < optSeOption.Length; k++)
                                //{
                                //    MessageBox.Show(optSeOption[k].Trim());
                                //    if (PartOptTen.Trim() == optSeOption[k].Trim() && PartOptTen.Trim() != "")//위의 로직대로 나누면 공백이 생길 수 있고 사용자가 실수로 공백을 넣을 수 있어 Trim으로 공백제거
                                //    {
                                //        totalSeq = Services.ApplicationServices.StructureSvc.GetNewSequence(entityId, 1, 1);
                                //        Services.ApplicationServices.StructureSvc.IsExistStructure(entityId, "EBOM", -1, Convert.ToString(datarow["child_id"])); //쿼리로 변경 해야 함 시퀀스까지 넣으면 비교가 안됨.
                                //        //MessageBox.Show("구조 체크 나간 후");
                                //        Services.ApplicationServices.StructureSvc.CreateStructure(entityId, "EBOM", userid, totalSeq, Convert.ToString(datarow["child_id"]), 1);

                                //    }
                                //}
                                //optSeOption.Initialize();// 배열사용에 따른 정보가 남을 수 있어 초기화 함.
                                //dataRowOption = string.Empty;

                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace + ", " + ex.Message);
            }
        }
        public string PMEcoValidation(string orgId,string newId, string userId)
        {

          
     
            string AffectedProduct = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$AffectedProduct");//설계 변경 통보의 번호        
            string AffectedRegul = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$AffectedRegul");//설계 변경 통보의 번호
            string ReasonChange = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$ReasonChange");//설계 변경 통보의 번호
            string ProductCorrection = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$ProductCorrection");//설계 변경 통보의 번호
            string SupplierImprove = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$SupplierImprove");//설계 변경 통보의 번호
            string CostReduction = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$CostReduction");//설계 변경 통보의 번호
            string ManufactureImprove = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$ManufactureImprove");//설계 변경 통보의 번호
            string ScopeChange = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$ScopeChange");//설계 변경 통보의 번호
            string QualityImprove = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$QualityImprove");//설계 변경 통보의 번호        
            string InitialReleaser = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$InitialReleaser");//설계 변경 통보의 번호
            string RecordChange = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$RecordChange");//설계 변경 통보의 번호
            string Other = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Other");//설계 변경 통보의 번호
            string ReCustomer = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$ReCustomer");//설계 변경 통보의 번호        
            string Manufacturing = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Manufacturing");//설계 변경 통보의 번호
            string Supplier = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Supplier");//설계 변경 통보의 번호
            string ReDesign = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$ReDesign");//설계 변경 통보의 번호
            string Prog = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Prog");//설계 변경 통보의 번호        
            string Tooling = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Tooling");//설계 변경 통보의 번호        
            string Engineering = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Engineering");//설계 변경 통보의 번호        
            string Purchasing = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Purchasing");//설계 변경 통보의 번호        
            string Quality = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Quality");//설계 변경 통보의 번호        
            string OtherTwo = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$OtherTwo");//설계 변경 통보의 번호        
            string Urgent = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Urgent");//설계 변경 통보의 번호        
            string Proto = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Proto");//설계 변경 통보의 번호        
            string SOP = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$SOP");//설계 변경 통보의 번호        
            string Running = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$Running");//설계 변경 통보의 번호        
            string OtherThr = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$OtherThr");//설계 변경 통보의 번호        
            string UrgentDate = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$UrgentDate");//설계 변경 통보의 번호        
            string ProtoDate = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$ProtoDate");//설계 변경 통보의 번호        
            string SOPDate = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$SOPDate");//설계 변경 통보의 번호    
            string RunningDate = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$RunningDate");//설계 변경 통보의 번호        
            string ChangeDes = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$ChangeDes");//설계 변경 통보의 번호        
            string CostImpact = Services.ApplicationServices.ObjectSvc.GetPropertyValue(newId, "E$CostImpact");//설계 변경 통보의 번호            

            int i = 0;
            int j = 0;
            int k = 0;
       

            if (AffectedProduct.Length < 3)
            {
                //MessageBox.Show(AffectedProduct + " AffectedProduct는 3글자 이상이어야 합니다.");
                return Adaptive.Action.ActionResultMessage.GetXmlMessage(Adaptive.Action.ActionResultType.Error, "Affected Product(Model) List는 3글자 이상이어야 합니다.");

            }

            if (AffectedRegul.Length < 3)
            {
                return Adaptive.Action.ActionResultMessage.GetXmlMessage(Adaptive.Action.ActionResultType.Error, "Affected Regulatory(Certification) List는 3글자 이상이어야 합니다.");

            }

            if (ReasonChange.Length < 10)
            {
                return Adaptive.Action.ActionResultMessage.GetXmlMessage(Adaptive.Action.ActionResultType.Error, "Reason for Change는 10글자 이상이어야 합니다.");
            }

            if (ChangeDes.Length < 20)
            {
                return Adaptive.Action.ActionResultMessage.GetXmlMessage(Adaptive.Action.ActionResultType.Error, "Change Description는 20글자 이상이어야 합니다.");
            }

            if (CostImpact.Length < 10)
            {
                return Adaptive.Action.ActionResultMessage.GetXmlMessage(Adaptive.Action.ActionResultType.Error, "Cost Impact Information는 10글자 이상이어야 합니다.");
            }




            if (ProductCorrection == "True")
            {
                i++;

            }

            if (SupplierImprove == "True")
            {
                i++;
            }

            if (CostReduction == "True")
            {
                i++;

            }

            if (ManufactureImprove == "True")
            {
                i++;

            }

            if (ScopeChange == "True")
            {
                i++;

            }

            if (QualityImprove == "True")
            {
                i++;

            }

            if (InitialReleaser == "True")
            {
                i++;

            }

            if (RecordChange == "True")
            {
                i++;

            }

            if (Other == "True")
            {
                i++;

            }

            if (i != 1)
            {
                //MessageBox.Show(" Category of Changed는 하나만 선택 가능 합니다.");
                return Adaptive.Action.ActionResultMessage.GetXmlMessage(Adaptive.Action.ActionResultType.Error, "Category of Changed는 하나만 선택 가능 합니다.");
            }


            if (Manufacturing == "True")
            {
                j++;

            }

            if (Supplier == "True")
            {
                j++;

            }

            if (ReCustomer == "True")
            {
                j++;

            }

            if (Prog == "True")
            {
                j++; 

            }

            if (Tooling == "True")
            {
                j++;
            }

            if (ReDesign == "True")
            {
                j++;
            }

            if (Purchasing == "True")
            {
                j++;
            }
            if (Engineering == "True")
            {
                j++;
            }

            if (Quality == "True")
            {
                j++;
            }

            if (OtherTwo == "True")
            {
                j++;
            }
            if (j != 1)
            {
                //MessageBox.Show(" Reported By는 하나만 선택 가능 합니다.");
                return Adaptive.Action.ActionResultMessage.GetXmlMessage(Adaptive.Action.ActionResultType.Error, "Reported By는 하나만 선택 가능 합니다.");
            }

            if (Urgent == "True")
            {
                k++;

            }

            if (Proto == "True")
            {
                k++;
            }

            if (SOP == "True")
            {
                k++;
            }

            if (Running == "True")
            {
                k++;

            }

            if (OtherThr == "True")
            {
                k++;

            }

            if (k != 1)
            {
                //MessageBox.Show(" Proposed Implementation Date는 하나만 선택 가능 합니다.");
                return Adaptive.Action.ActionResultMessage.GetXmlMessage(Adaptive.Action.ActionResultType.Error, "Proposed Implementation Date는 하나만 선택 가능 합니다.");
            }

            return "True" ;

        }

        public void pmmbom(string entityId, string userid)
        {

            string topPartNumber = string.Empty;
            string partNumber = string.Empty;
           
            try
            {

                XmlDocument doc = new XmlDocument();

                string templatePath = Configurations.StartupPath + "\\PMBOM.xls";               

                C1XLBook book = new C1XLBook();
                book.Load(templatePath);

                XLSheet sheet;
                sheet = book.Sheets[0];

                XLCell cell;
                
              
                // 분류의 기종 row 색상
                XLStyle styleTOPone = new XLStyle(book);
                styleTOPone.AlignHorz = XLAlignHorzEnum.Center;
                styleTOPone.AlignVert = XLAlignVertEnum.Center;
                styleTOPone.Font = new Font("돋움", 9, FontStyle.Regular);
                styleTOPone.BorderBottom = XLLineStyleEnum.Thin;
                styleTOPone.BorderLeft = XLLineStyleEnum.Thin;
                styleTOPone.BorderRight = XLLineStyleEnum.Thin;
                styleTOPone.BorderTop = XLLineStyleEnum.Thin;
                //styleTOPone.BackColor = Color.FromArgb(120, 100, 240);
                styleTOPone.WordWrap = true;
                styleTOPone.Locked = false;

                topPartNumber = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Number");//설계 변경 통보의 번호


                if (topPartNumber.Length > 11)
                {
                    cell = sheet[1, 0];
                    cell.Value = topPartNumber;

                }
                else
                {
                    topPartNumber = topPartNumber + "-" + Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "Version");
                    cell = sheet[1, 0];
                    cell.Value = topPartNumber;

                }


                cell = sheet[1, 1];
                cell.Value = 0;

                cell = sheet[1, 2];
                cell.Value = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Name");
                string modelName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$Modelname"); 
                cell = sheet[1, 4];
                cell.Value = modelName;

                cell = sheet[1, 5];
                cell.Value = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "P$DESC");

                sheet.Name = modelName;
                for (int k = 0; k < 18; k++)
                {


                    sheet[1, k].Style = styleTOPone;

                }

                int i = 1;
                int j = 0;
                int bomRow = 2;
                string[] berforetreeseq = new string[100];
                int[] level = new int[20];   
                int berforetreelevel = 0;              
             
                for (int l = 0; l < 20; l++)
                {
                    
                    level[l] = 0;

                }


                cell = sheet[1, 1];
                cell.Value = "0";

                 DataTable rowDt = Services.ApplicationServices.StructureSvc.GetTopDownStructure(entityId,"EBOM","part");
               
                 foreach (DataRow datarow in rowDt.Rows)
                 {

                     partNumber = Convert.ToString(datarow["P$Number"]);

                     if (partNumber.Length > 11)
                     {
                         cell = sheet[bomRow, 0];
                         cell.Value = partNumber;

                     }
                     else
                     {   
                         cell = sheet[bomRow, 0];
                         cell.Value = partNumber + "-" + Convert.ToString(datarow["P$Version"]); 

                     }

                     int treelevel = Convert.ToInt32(datarow["tree_level"]);

                     if (berforetreelevel < treelevel)
                     {

                         if (treelevel == 1)
                         {
                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow, 1];
                             cell.Value = level[treelevel];
                             

                         }

                         if (treelevel == 2)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow - 1, 1];
                             cell.Value = level[treelevel];

                         }

                         if (treelevel == 3)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }

                         if (treelevel == 4)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 2] + "-" + level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }

                         if (treelevel == 5)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 3] + "-" + level[treelevel - 2] + "-" + level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }

                         if (treelevel == 6)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 4] + "-" + level[treelevel - 3] + "-" + level[treelevel - 2] + "-" + level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }

                         if (treelevel == 7)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 5] + "-" + level[treelevel - 4] + "-" + level[treelevel - 3] + "-" + level[treelevel - 2] + "-" + level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }

                         if (treelevel == 8)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 6] + "-" + level[treelevel - 5] + "-" + level[treelevel - 4] + "-" + level[treelevel - 3] + "-" + level[treelevel - 2] + "-" + level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }

                         if (treelevel == 9)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 7] + "-" + level[treelevel - 6] + "-" + level[treelevel - 5] + "-" + level[treelevel - 4] + "-" + level[treelevel - 3] + "-" + level[treelevel - 2] + "-" + level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }

                         if (treelevel == 10)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 8] + "-" + level[treelevel - 7] + "-" + level[treelevel - 6] + "-" + level[treelevel - 5] + "-" + level[treelevel - 4] + "-" + level[treelevel - 3] + "-" + level[treelevel - 2] + "-" + level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }

                         if (treelevel == 11)
                         {

                             level[treelevel] = level[treelevel] + 1;
                             cell = sheet[bomRow -1, 1];
                             cell.Value = level[treelevel - 9] + "-" + level[treelevel - 8] + "-" + level[treelevel - 7] + "-" + level[treelevel - 6] + "-" + level[treelevel - 5] + "-" + level[treelevel - 4] + "-" + level[treelevel - 3] + "-" + level[treelevel - 2] + "-" + level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]); ;

                         }


                         //if (treelevel == 1)
                         //{
                         //    level[treelevel] = level[treelevel] + 1;
                         //    cell = sheet[bomRow, 2];
                         //    cell.Value = level[treelevel];
                            
                         //}
                         //else
                         //{
                         //    if (treelevel > 2)
                         //    {
                         //        berforetreeseq[treelevel] = berforetreeseq[treelevel] + 1;
                         //        level[treelevel] = level[treelevel - 1] + "-" + Convert.ToString(level[treelevel]);
                         //        cell = sheet[bomRow - 1, 2];
                         //        cell.Value = level[treelevel];
                         //        //berforetreeseq[treelevel + 1] = 0;
                         //    }
                         //}
                         
                         berforetreelevel = treelevel;                 

                        
                     }
                     else
                     {
                         if (berforetreelevel > treelevel)
                         {
                             //if (treelevel == 1)
                             //{

                             //    //level[treelevel] = level[treelevel] + 1;
                             //    //cell = sheet[bomRow, 2];
                             //    //cell.Value = level[treelevel];
                             //    //berforetreeseq[treelevel] = Convert.ToString(level[treelevel]);

                             //}
                             //else
                             //{

                             //    //if (berforetreelevel - treelevel > 1)
                             //    //{
                             //    //    for (int l = treelevel + 1; l < 20; l++)
                             //    //    {

                             //    //        level[l] = 0;

                             //    //    }
                             //    //}

                             //    //berforetreeseq[treelevel] = berforetreeseq[treelevel - 1] + "-" + Convert.ToString(level[treelevel]);
                             //    //cell = sheet[bomRow, 2];
                             //    //cell.Value = berforetreeseq[treelevel];


                             //}

                             for (int l = treelevel + 2; l < 20; l++)
                             {

                                 level[l] = 0;

                             }
                           
                             berforetreelevel = treelevel;
                         }
                         //if (treelevel < berforetreelevel)
                         //{
                         //    if (treelevel == 1)
                         //    {

                         //        berforetreeseq[treelevel] = berforetreeseq[treelevel] + 1;
                         //        level[treelevel] = Convert.ToString(berforetreeseq[treelevel]);
                         //        level[treelevel + 1] = Convert.ToString(berforetreeseq[treelevel]);
                         //        cell = sheet[bomRow, 2];
                         //        cell.Value = level[treelevel];
                         //        berforetreeseq[treelevel + 2] = 0;
                         //    }
                         //    else
                         //    {
                         //        if (treelevel >= 3)
                         //        {

                         //            berforetreeseq[treelevel] = berforetreeseq[treelevel] + 1;
                         //            level[treelevel] = level[treelevel - 1] + "-" + Convert.ToString(berforetreeseq[treelevel]);
                         //            cell = sheet[bomRow - 1, 2];
                         //            cell.Value = level[treelevel];

                         //        }
                         //    }
                         //}
                         //else
                         //{

                         //    if (treelevel == 1)
                         //    {

                         //        berforetreeseq[treelevel] = berforetreeseq[treelevel] + 1;
                         //        level[treelevel] = Convert.ToString(berforetreeseq[treelevel]);
                         //        level[treelevel + 1] = Convert.ToString(berforetreeseq[treelevel]);
                         //        cell = sheet[bomRow, 2];
                         //        cell.Value = level[treelevel];
                         //        berforetreeseq[treelevel + 2] = 0;
                         //    }
                         //    else
                         //    {
                         //        //berforetreeseq[treelevel] = berforetreeseq[treelevel] + 1;
                         //        //level[treelevel] = level[treelevel - 1] + "-" + Convert.ToString(berforetreeseq[treelevel]);
                         //        //cell = sheet[bomRow - 1, 0];
                         //        //cell.Value = level[treelevel];

                               
                         //    }
                         //}

                         //berforetreelevel = treelevel;
                         //MessageBox.Show(level[treelevel -1] + ";" + level[treelevel] + ";" + level[treelevel + 1]);
                         //같은 레벨의 부품
                     }



                     cell = sheet[bomRow, 2];
                     cell.Value = Convert.ToString(datarow["P$Name"]); ;


                     cell = sheet[bomRow, 3];
                     cell.Value = Convert.ToString(datarow["P$Detailname"]); 

                     cell = sheet[bomRow, 4];
                     cell.Value = Convert.ToString(datarow["P$Modelname"]);

                     cell = sheet[bomRow, 5];
                     string descP = Convert.ToString(datarow["P$DESC"]);
                     if (Convert.ToString(datarow["P$DESC"]).Length > 45)
                     {
                          descP = Convert.ToString(datarow["P$DESC"]).Insert(44, "\n");
                         if (Convert.ToString(datarow["P$DESC"]).Length > 90)
                         {
                             descP = descP.Insert(89, "\n");
                         }
                     }
                     cell.Value = descP;

                     cell = sheet[bomRow, 6];
                     cell.Value = Convert.ToString(datarow["Quantity"]); //수량

                     cell = sheet[bomRow, 7];

                     string partLocation = Convert.ToString(datarow["S$Location"]); //위치
                     if (Convert.ToString(datarow["S$Location"]).Length > 45)
                     {
                         partLocation = Convert.ToString(datarow["S$Location"]).Insert(44, "\n");
                         if (Convert.ToString(datarow["S$Location"]).Length > 90)
                         {
                             partLocation = partLocation.Insert(89, "\n");
                         }
                     }
                     cell.Value = partLocation;

                     cell = sheet[bomRow, 8];
                     cell.Value = Convert.ToString(datarow["S$AlternativeOrder"]); //대체순번

                     cell = sheet[bomRow, 9];
                     cell.Value = Convert.ToString(datarow["P$Marker"]); //maker

                     cell = sheet[bomRow, 10];
                     cell.Value = Convert.ToString(datarow["S$Remark"]); //remark

                     cell = sheet[bomRow, 11];
                     cell.Value = Convert.ToString(datarow["P$UL"]); //remark

                     cell = sheet[bomRow, 12];
                     cell.Value = Convert.ToString(datarow["P$ROHS"]); //remark

                     cell = sheet[bomRow, 13];
                     cell.Value = Convert.ToString(datarow["P$MSDS"]); //remark


                     cell = sheet[bomRow, 14];
                     cell.Value = Convert.ToString(datarow["P$REACH"]); //remark

                     cell = sheet[bomRow, 15];
                     cell.Value = Convert.ToString(datarow["P$ApprovalSheet"]); //remark

                     cell = sheet[bomRow, 16];
                     cell.Value = Convert.ToString(datarow["P$Datasheet"]); //remark


                     for (int k = 0; k < 17; k++)
                     {








                         sheet[bomRow, k].Style = styleTOPone;
                        
                     }


                     bomRow++;
                     
                    
                 }

                //string prevId = Services.ApplicationServices.VersioningSvc.GetPreviousID(entityId);

                //DataTable prevDt = Services.ApplicationServices.StructureSvc.GetTopDownStructure("PartStructure", "Part", prevId);
                ////MessageBox.Show("이전ID");








                 string exportPath = Configurations.WorkspacePath + "\\" + modelName + "_" + topPartNumber + "_BOM.xls";

                

                



                ////MessageBox.Show("1");
                //DataTable MBOMDt = Services.ApplicationServices.StructureSvc.GetTopDownStructure("PartStructure", "Part", entityId );

                //// 선택된 오브젝트의 파트리스트를 엑셀에 쓰는 로직
               

                //    //MessageBox.Show("1");
                //    sheet.Rows[sheetRow].Height = 500;

                //    sheetRow++;



                    book.Save(exportPath);

                    //"지금 어플리케이션으로 내용을 확인하시겠습니까?"
                    if (MessageBox.Show(Services.NativeSvc.GetNativeMessage("AreyouGoingToCheckTheContentsWithTheApplicationNow"), Services.NativeSvc.GetNativeCaption("Export"), MessageBoxButtons.OKCancel) == DialogResult.OK)
                        System.Diagnostics.Process.Start(exportPath);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace + ", " + ex.Message);
            }
           
        }
        public void PmEco(string entityId, string userid)
        {
            List<string> downloadList = new List<string>();

           
            C1XLBook book = new C1XLBook();

            XmlDocument doc = new XmlDocument();

            XLSheet sheet, sheet1;

            XLCell cell;
            XLRow row = new XLRow();

            XLStyle stylerow = new XLStyle(book);

            
            XLStyle styleleft = new XLStyle(book);
            styleleft.AlignHorz = XLAlignHorzEnum.Center;
            styleleft.AlignVert = XLAlignVertEnum.Center;
            styleleft.Font = new Font("돋움", 10);
            styleleft.BorderBottom = XLLineStyleEnum.Thin;
            styleleft.BorderLeft = XLLineStyleEnum.Thick;
            styleleft.BorderRight = XLLineStyleEnum.Thin;
            styleleft.BorderTop = XLLineStyleEnum.Thin;
            //styleTOPone.BackColor = Color.FromArgb(120, 100, 240);
            styleleft.WordWrap = true;
            styleleft.Locked = true;

            XLStyle styleRight = new XLStyle(book);
            styleRight.AlignHorz = XLAlignHorzEnum.Center;
            styleRight.AlignVert = XLAlignVertEnum.Center;
            styleRight.Font = new Font("돋움", 10);
            styleRight.BorderBottom = XLLineStyleEnum.Thin;
            styleRight.BorderLeft = XLLineStyleEnum.Thin;
            styleRight.BorderRight = XLLineStyleEnum.Thick;
            styleRight.BorderTop = XLLineStyleEnum.Thin;
            //styleTOPone.BackColor = Color.FromArgb(120, 100, 240);
            styleRight.WordWrap = true;
            styleRight.Locked = true;

            XLStyle stylecenter = new XLStyle(book);
            stylecenter.AlignHorz = XLAlignHorzEnum.Center;
            stylecenter.AlignVert = XLAlignVertEnum.Center;
            stylecenter.Font = new Font("돋움", 10);
            stylecenter.BorderBottom = XLLineStyleEnum.Thin;
            stylecenter.BorderLeft = XLLineStyleEnum.Thin;
            stylecenter.BorderRight = XLLineStyleEnum.Thin;
            stylecenter.BorderTop = XLLineStyleEnum.Thin;
            //styleTOPone.BackColor = Color.FromArgb(120, 100, 240);
            stylecenter.WordWrap = true;
            stylecenter.Locked = true;
           
            Adaptive.Archive.DownloadGate downSvc = new Adaptive.Archive.DownloadGate();
           

            //Waiting.Show();

            try
            {
                //string templatePathA = "C:\\Users\\기환\\Desktop\\AoneNX\\Workspace"  + "\\PmEco2.xls";//설계 변경 통보의 기본 양식을 불러옴
                string templatePathA = Configurations.StartupPath + "\\PmEco2.xls";
                //MessageBox.Show(templatePathA);
                book.Load(templatePathA);

            }
            catch
            {
                MessageBox.Show("설계변경통보서의 기본양식 파일이 없습니다.");
            }

            try
            {

                string ecoNumber = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Number");//설계 변경 통보의 번호
                string DateIntiated = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$DateIntiated");//설계 변경 통보의 번호
                string Product = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Product");//설계 변경 통보의 번호
                string Customer = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Customer");//설계 변경 통보의 번호
                string AffectedProduct = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$AffectedProduct");//설계 변경 통보의 번호        
                string AffectedRegul = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$AffectedRegul");//설계 변경 통보의 번호
                string ReasonChange = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$ReasonChange");//설계 변경 통보의 번호
                string ProductCorrection = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$ProductCorrection");//설계 변경 통보의 번호
                string SupplierImprove = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$SupplierImprove");//설계 변경 통보의 번호
                string CostReduction = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$CostReduction");//설계 변경 통보의 번호
                string ManufactureImprove = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$ManufactureImprove");//설계 변경 통보의 번호
                string ScopeChange = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$ScopeChange");//설계 변경 통보의 번호
                string QualityImprove = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$QualityImprove");//설계 변경 통보의 번호        
                string InitialReleaser = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$InitialReleaser");//설계 변경 통보의 번호
                string RecordChange = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$RecordChange");//설계 변경 통보의 번호
                string Other = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Other");//설계 변경 통보의 번호
                string ReCustomer = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$ReCustomer");//설계 변경 통보의 번호        
                string Manufacturing = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Manufacturing");//설계 변경 통보의 번호
                string Supplier = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Supplier");//설계 변경 통보의 번호
                string ReDesign = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$ReDesign");//설계 변경 통보의 번호
                string Prog = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Prog");//설계 변경 통보의 번호        
                string Tooling = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Tooling");//설계 변경 통보의 번호        
                string Engineering = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Engineering");//설계 변경 통보의 번호        
                string Purchasing = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Purchasing");//설계 변경 통보의 번호        
                string Quality = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Quality");//설계 변경 통보의 번호        
                string OtherTwo = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$OtherTwo");//설계 변경 통보의 번호        
                string Urgent = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Urgent");//설계 변경 통보의 번호        
                string Proto = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Proto");//설계 변경 통보의 번호        
                string SOP = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$SOP");//설계 변경 통보의 번호        
                string Running = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$Running");//설계 변경 통보의 번호        
                string OtherThr = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$OtherThr");//설계 변경 통보의 번호        
                string UrgentDate = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$UrgentDate");//설계 변경 통보의 번호        
                string ProtoDate = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$ProtoDate");//설계 변경 통보의 번호        
                string SOPDate = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$SOPDate");//설계 변경 통보의 번호    
                string RunningDate = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$RunningDate");//설계 변경 통보의 번호        
                string ChangeDes = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$ChangeDes");//설계 변경 통보의 번호        
                string CostImpact = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$CostImpact");//설계 변경 통보의 번호            
                string printUser = Convert.ToString(Services.ApplicationServices.SecuritySvc.ServerTime().Year) + "-" + Convert.ToString(Services.ApplicationServices.SecuritySvc.ServerTime().Month) + "-" + Convert.ToString(Services.ApplicationServices.SecuritySvc.ServerTime().Day);//출력일자 
                string createUser = Services.ApplicationServices.ObjectSvc.GetCreator(entityId);
                string createId = Services.ApplicationServices.ObjectSvc.GetCreatorID(entityId);
                string OtherContent = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$OtherContent");//설계 변경 통보의 번호        
                string OtherTwoContent = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$OtherTwoContent");//설계 변경 통보의 번호        
                string OtherThrContent = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$OtherThrContent");//설계 변경 통보의 번호    
                string department = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "E$department");
                AffectedProduct = AffectedProduct.Replace("&CrLf", "\r\n");
                AffectedRegul = AffectedRegul.Replace("&CrLf", "\r\n");
                ReasonChange = ReasonChange.Replace("&CrLf", "\r\n");
                Other = Other.Replace("&CrLf", "\r\n");
                OtherTwo = OtherTwo.Replace("&CrLf", "\r\n");
                ChangeDes = ChangeDes.Replace("&CrLf", "\r\n");
                CostImpact = CostImpact.Replace("&CrLf", "\r\n");

                sheet = book.Sheets[0];
                int l = entityId.IndexOf('.');
                string engClassname = entityId.Substring(0, l);
                //string engClassname2 = Services.ApplicationServices.ObjectSvc.GetPropertyValue(entityId, "C$Name");//설계 변경 통보의 번호
                //MessageBox.Show(engClassname);

                if (engClassname == "E$PMEC-1049")
                {
                    cell = sheet[0, 7];
                    cell.Value = "Engineering Change Order";

                    cell = sheet[0, 33];
                    cell.Value = "ECO";
                }

                if (engClassname == "E$PMEC-1048")
                {
                    cell = sheet[0, 7];
                    cell.Value = "Engineering Change Request";

                    cell = sheet[0, 33];
                    cell.Value = "ECR";
                }

                if (engClassname == "E$PMEC-1064")
                {
                    cell = sheet[0, 7];
                    cell.Value = "Engineering Change Notice";

                    cell = sheet[0, 33];
                    cell.Value = "ECN";
                }





                //MessageBox.Show("2");
                //DataTable prevDt = Services.ApplicationServices.MembershipSvc.GetAssignGroups(createId);




                //foreach (DataRow frdr in prevDt.Rows)
                //  {

                //    MessageBox.Show("1");

                //  }

                sheet.Name = ecoNumber;

                cell = sheet[0, 34];
                cell.Value = ecoNumber;





                //sheet.Locked = true;

                // 시방서의 정보를 셀단위로 엑셀에 삽입
                cell = sheet[1, 34];
                DateIntiated = DateIntiated.Substring(0, 10);
                cell.Value = DateIntiated;

                cell = sheet[3, 7];
                cell.Value = createUser;

                cell = sheet[3, 14];
                cell.Value = department;

                cell = sheet[3, 22];
                cell.Value = Product;

                cell = sheet[3, 32];
                cell.Value = Customer;

                cell = sheet[6, 2];
                cell.Value = AffectedProduct;

                cell = sheet[6, 21];
                cell.Value = AffectedRegul;

                cell = sheet[12, 2];
                cell.Value = ReasonChange;


                if (ProductCorrection == "True")
                {
                    cell = sheet[17, 10];
                    cell.Value = "√";

                }

                if (SupplierImprove == "True")
                {
                    cell = sheet[17, 21];
                    cell.Value = "√";

                }

                if (CostReduction == "True")
                {
                    cell = sheet[19, 2];
                    cell.Value = "√";

                }

                if (ManufactureImprove == "True")
                {
                    cell = sheet[19, 10];
                    cell.Value = "√";

                }

                if (ScopeChange == "True")
                {
                    cell = sheet[21, 2];
                    cell.Value = "√";

                }

                if (QualityImprove == "True")
                {
                    cell = sheet[21, 10];
                    cell.Value = "√";

                }

                if (InitialReleaser == "True")
                {
                    cell = sheet[23, 2];
                    cell.Value = "√";

                }

                if (RecordChange == "True")
                {
                    cell = sheet[23, 10];
                    cell.Value = "√";

                }

                if (Other == "True")
                {

                    cell = sheet[19, 21];
                    cell.Value = "√";
                    cell = sheet[19, 32];
                    cell.Value = OtherContent;
                }

                if (Manufacturing == "True")
                {
                    cell = sheet[26, 10];
                    cell.Value = "√";

                }

                if (Supplier == "True")
                {
                    cell = sheet[26, 21];
                    cell.Value = "√";

                }

                if (ReCustomer == "True")
                {
                    cell = sheet[28, 2];
                    cell.Value = "√";

                }

                if (Prog == "True")
                {
                    cell = sheet[28, 10];
                    cell.Value = "√";

                }

                if (Tooling == "True")
                {
                    cell = sheet[28, 21];
                    cell.Value = "√";

                }

                if (ReDesign == "True")
                {
                    cell = sheet[30, 2];
                    cell.Value = "√";

                }

                if (Purchasing == "True")
                {
                    cell = sheet[30, 10];
                    cell.Value = "√";

                }
                if (Engineering == "True")
                {
                    cell = sheet[32, 2];
                    cell.Value = "√";

                }

                if (Quality == "True")
                {
                    cell = sheet[32, 10];
                    cell.Value = "√";

                }

                if (OtherTwo == "True")
                {
                    cell = sheet[30, 21];
                    cell.Value = "√";

                    cell = sheet[30, 32];
                    cell.Value = OtherTwoContent;
                }

                if (Urgent == "True")
                {
                    cell = sheet[35, 10];
                    cell.Value = "√";

                    cell = sheet[35, 14];
                    cell.Value = UrgentDate;

                }

                if (Proto == "True")
                {
                    cell = sheet[35, 21];
                    cell.Value = "√";

                    cell = sheet[35, 27];
                    cell.Value = ProtoDate;

                }

                if (SOP == "True")
                {
                    cell = sheet[37, 2];
                    cell.Value = "√";



                }

                if (Running == "True")
                {
                    cell = sheet[37, 10];
                    cell.Value = "√";
                    cell = sheet[37, 14];
                    cell.Value = RunningDate;


                }

                if (OtherThr == "True")
                {
                    cell = sheet[37, 21];
                    cell.Value = "√";

                    cell = sheet[37, 33];
                    cell.Value = OtherThrContent;


                }


                cell = sheet[37, 6];
                SOPDate = SOPDate.Substring(0, 10);
                cell.Value = SOPDate;

                cell = sheet[40, 2];
                cell.Value = ChangeDes;

                cell = sheet[47, 2];
                cell.Value = CostImpact;



                sheet1 = book.Sheets[1];
                DataTable ECORelationdt = Services.ApplicationServices.DataSvc.ExecuteDataTable("select*, p.state as pstate from structure_info as s INNER JOIN  part_info as p on s.parent_id = p.urid   where s.S$ECOTopId = '" + entityId + "' AND s.S$ECOACTION is not null");

                //DataTable ECORelationdt = Services.ApplicationServices.RelationSvc.GetRelationList("EngineeringChange", entityId, "LinkedECO");

                int i = 1;

                foreach (DataRow ECOdr in ECORelationdt.Rows)
                {
                    string Classification = Convert.ToString(ECOdr["S$ECOACTION"]);//설계 변경 통보의 번호
                    string ECOReason = Convert.ToString(ECOdr["S$REASON"]);//설계 변경 통보의 번호
                    string ECOQtyOne = Convert.ToString(ECOdr["quantity"]);//설계 변경 통보의 번호
                    string ECOQtyTwo = Convert.ToString(ECOdr["S$ECOQuentity"]);//설계 변경 통보의 번호       
                    string ECOAFTERNUM = Convert.ToString(ECOdr["S$ECOAFTERNUM"]);//설계 변경 통보의 번호       
                    string ReferenceNo = Convert.ToString(ECOdr["S$ECOALTERNATE"]);//설계 변경 통보의 번호
                    string APPLIEDASSY = Services.ApplicationServices.ObjectSvc.GetPropertyValue(Convert.ToString(ECOdr["parent_id"]), "Number"); //설계 변경 통보의 번호
                    string APPLIEDPart = Services.ApplicationServices.ObjectSvc.GetPropertyValue(Convert.ToString(ECOdr["child_id"]), "Number");//설계 변경 통보의 번호
                    string StockManage = Convert.ToString(ECOdr["S$STOCK"]);//설계 변경 통보의 번호

                    //string DeletedParts = Convert.ToString(ECOdr["child_id"]);//설계 변경 통보의 번호
                    //DeletedParts = DeletedParts.Replace("&CrLf", "\n");




                    //sheet1.Rows.Insert(9);
                    int j = i + 8;

                    //sheet1.Rows[j].Style.WordWrap = true;
                    sheet1.Rows[j].Height = 1000;

                    cell = sheet1[1, 2];
                    cell.Value = ecoNumber;


                    cell = sheet1[j, 1];
                    cell.Value = Convert.ToString(i);
                    //cell.Style = styleleft;

                    cell = sheet1[j, 2];
                    cell.Value = ReferenceNo;
                    //cell.Style = stylecenter;

                    cell = sheet1[j, 3];
                    cell.Value = Classification;
                    //cell.Style = stylecenter;

                    cell = sheet1[j, 4];
                    cell.Value = ECOReason;
                    //cell.Style = stylecenter;

                    cell = sheet1[j, 5];
                    cell.Value = APPLIEDASSY;
                    //cell.Style = stylecenter;

                    cell = sheet1[j, 6];
                    cell.Value = ECOQtyOne;
                    //cell.Style = stylecenter;


                    //cell.Style = stylecenter;

                    cell = sheet1[j, 8];
                    cell.Value = StockManage;

                    cell = sheet1[j, 10];
                    cell.Value = ECOQtyTwo;
                    //cell.Style = stylecenter;



                    if (Classification == "ADD")
                    {
                        cell = sheet1[j, 7];
                        cell.Value = APPLIEDPart;

                    }

                    if (Classification == "ALT")
                    {
                        cell = sheet1[j, 7];
                        cell.Value = APPLIEDPart;

                    }

                    if (Classification == "CHG")
                    {
                        cell = sheet1[j, 7];
                        cell.Value = ECOAFTERNUM;

                        cell = sheet1[j, 11];
                        cell.Value = APPLIEDPart;

                    }

                    if (Classification == "DEL")
                    {
                        cell = sheet1[j, 11];
                        cell.Value = APPLIEDPart;

                    }


                    i++;
                }

                //결재 정보 로직


                DataTable engChangeWorkDt = Services.ApplicationServices.DataSvc.ExecuteDataTable("SELECT * FROM work_routing_info where master_id  = '" + entityId + "'"); //상위관계 카운트    






                foreach (DataRow workrow in engChangeWorkDt.Rows)
                {

                    try
                    {
                        if (Convert.ToString(workrow["state_name"]) == "Start")
                        {

                            cell = sheet[54, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);



                            if (Convert.ToString(workrow["close_date"]) == "NULL")
                            {

                            }
                            else
                            {

                                cell = sheet[54, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                //XLPictureShape pic = new XLPictureShape(JPG,-1000,-1000);
                                //pic.LineColor = Color.Black ;
                                //pic.LineWidth = 100;



                                sheet[53, 26].Value = JPG;
                                downloadList.Add(downFileid);



                                //cell = sheet[53, 19];
                                //cell.Value = Convert.ToString(workrow["comments"]);


                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "RnD Reviewer")
                        {

                            cell = sheet[57, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);




                            if (Convert.ToString(workrow["close_date"]) == "NULL")
                            {

                            }
                            else
                            {
                                cell = sheet[57, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[56, 26].Value = JPG;
                                downloadList.Add(downFileid);








                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "RnD Approval")
                        {

                            cell = sheet[60, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            //MessageBox.Show(Convert.ToString(workrow["close_date"]));
                            if (Convert.ToString(workrow["close_date"]) == "NULL")
                            {
                            }
                            else
                            {
                                cell = sheet[60, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[59, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "RnD 유관부서1")
                        {

                            cell = sheet[68, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);


                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                ////sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[68, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[67, 26].Value = JPG;
                                downloadList.Add(downFileid);




                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "RnD 유관부서2")
                        {

                            cell = sheet[71, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[71, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[70, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "Procurement")
                        {

                            cell = sheet[74, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[74, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[73, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "Quality")
                        {

                            cell = sheet[77, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[77, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[76, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "Marketing")
                        {

                            cell = sheet[80, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[80, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[79, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "Project Manager")
                        {

                            cell = sheet[83, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[83, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[82, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }


                        if (Convert.ToString(workrow["state_name"]) == "Material Controller")
                        {

                            cell = sheet[86, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[86, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[85, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "Production Controller")
                        {

                            cell = sheet[89, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[89, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[88, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "Production Engineering")
                        {

                            cell = sheet[92, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[92, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[91, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }

                        if (Convert.ToString(workrow["state_name"]) == "RnD Executive Director")
                        {

                            cell = sheet[63, 10];
                            cell.Value = Convert.ToString(workrow["owner"]);
                            if (Convert.ToString(workrow["close_date"]) == "")
                            {
                                //sheet.Rows.RemoveAt(workFlow);
                                //sheet.Rows.RemoveAt(workFlow);
                            }
                            else
                            {
                                cell = sheet[63, 34];
                                cell.Value = Convert.ToString(workrow["close_date"]).Substring(0, 10);

                                string downFileid = Services.ApplicationServices.RelationSvc.GetFromID("ATTACHEDFILE", Convert.ToString(workrow["owner_id"]));

                                Adaptive.Archive.DownloadGate dfd = new Adaptive.Archive.DownloadGate();

                                string signFileName = Services.ApplicationServices.ObjectSvc.GetPropertyValue(downFileid, "F$Name");//설계 변경 통보의 번호
                                string downPathJpg = Configurations.StartupPath + "\\" + "signimage";

                                dfd.Download(downFileid, downPathJpg);


                                Image JPG = Image.FromFile(downPathJpg + "\\" + signFileName);
                                XLPictureShape pic = new XLPictureShape(JPG, 700, 700);
                                pic.LineColor = Color.Black;
                                pic.LineWidth = 100;



                                sheet[62, 26].Value = JPG;
                                downloadList.Add(downFileid);

                            }

                        }




                    }
                    catch (Exception ex)
                    {
                        Adaptive.Log.WriteLine("RelECO", ex.Message);


                    }
                }
                //cell.Style = styleRight;
                //string fileDownPath = Services.ApplicationServices.VaultSvc.GetStorePath();
                //string relFileName = Convert.ToString(ECOdr["filename"]);
                //MessageBox.Show(fileDownPath);
                //int relFileNameLength = relFileName.Length - 3;
                ////string relFileNameLength1 = Convert.ToString(relFileNameLength);
                //string relFileNameSub = relFileName.Substring(relFileNameLength);
                //MessageBox.Show(relFileNameSub);

                //if (relFileNameSub == null)
                //{

                //}
                //else
                //{
                //    if (relFileNameSub.ToLower() == "jpg")
                //    {
                //        sheet1.Rows[j].Height = 2820;// 그림이 있는 경우 사용


                //        string relFileentityId = Convert.ToString(frdr["entityId"]);
                //        downSvc.Download(relFileentityId, fileDownPath);
                //        Image JPG = Image.FromFile("C:\\Users\\기환\\Desktop\\AoneNX\\Workspace"+ "\\" + relFileName);
                //        XLPictureShape pic = new XLPictureShape(JPG, 1500, 1500);
                //        pic.LineColor = Color.Black;
                //        pic.LineWidth = 100;


                //        sheet1[j, 7].Value = JPG;
                //        sheet1[j, 7].Style = styleTOPone;

                //    }

                //    if (relFileNameSub.ToLower() == "pdf")
                //    {
                //        string relFileentityId = Convert.ToString(frdr["entityId"]);
                //        downSvc.Download(relFileentityId, fileDownPath);
                //        cell = sheet1[j, 27];     
                //        cell.Value = "■";
                //        sheet1[j, 27].Hyperlink = "C:\\Users\\기환\\Desktop\\AoneNX\\Workspace" + "\\" + relFileName;
                //        sheet1[j, 27].Style = styleTOPone;

                //    }

                string exportPath = Configurations.WorkspacePath + "\\" + ecoNumber + ".xls";
                //MessageBox.Show(exportPath);
                try
                {
                    book.Save(exportPath);
                    if (MessageBox.Show(Services.NativeSvc.GetNativeMessage("AreyouGoingToCheckTheContentsWithTheApplicationNow"), Services.NativeSvc.GetNativeCaption("Export"), MessageBoxButtons.OKCancel) == DialogResult.OK)
                        System.Diagnostics.Process.Start(exportPath);





                }
                catch (Exception ex)
                {
                    MessageBox.Show("Workspace가 지정되어 있지 않거나, 같은 번호의 설계변경통보서가 열려 있습니다.");
                    Adaptive.Log.WriteLine("PMPrint", ex.Message);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace + ", " + ex.Message);
                Adaptive.Log.WriteLine("PMPrint", ex.Message);
            }


            foreach (string fileName in downloadList)
            {
                File.Delete(fileName);

            }
                   



                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("JVMPrint", ex.Message);
                //    Adaptive.Log.WriteLine("JVMPrint", ex.Message);
                //}



        






            //long childId = Convert.ToInt64(CurrentDr["id"].Value);
            //DataTable parentDt = Services.ApplicationServices.DataSvc.ExecuteDataTable("SELECT * FROM PartStructureInfo WHERE childId=" + childId); //상위관계 카운트




            //"지금 어플리케이션으로 내용을 확인하시겠습니까?"
                

        }
        public string topid = string.Empty;

        public void BOMecostart(string topid, string userid)
        {
            //DataTable startdt2 = Services.ApplicationServices.DataSvc.ExecuteDataTable("select*, p.state as pstate from structure_info as s INNER JOIN  part_info as p on s.parent_id = p.urid   where s.S$ECOTopId = '" + this.TopID + "' AND s.S$ECOACTION is null AND p.latest != '1'");
            // foreach (DataRow startdr2 in startdt2.Rows)
            // {
            //최신 버전이 아닌 것 선 체크하는 로직 추가 예정
            // }

            string checkcount = "1";
            DataTable startdt = Services.ApplicationServices.DataSvc.ExecuteDataTable("select*, p.state as pstate from structure_info as s INNER JOIN  part_info as p on s.parent_id = p.urid   where s.S$ECOTopId = '" + topid + "' AND s.S$ECOACTION is not null AND p.latest = '1'");
            Adaptive.Windows.Waiting.Show();
            foreach (DataRow startdr in startdt.Rows)
            {


                Application.DoEvents();
                string startParentId = Convert.ToString(startdr["parent_id"]);
                try
                {
                    string sqlStu2 = string.Format("select*, p.state as pstate from structure_info as s INNER JOIN  part_info as p on s.parent_id = p.urid where s.child_id = '" + startParentId + "' AND p.latest = '1'");
                    DataTable structureInfoDtStu2 = Services.ApplicationServices.DataSvc.ExecuteDataTable(sqlStu2);

                    string startParentnewId = string.Empty;
                    string startFromId = Convert.ToString(startdr["child_id"]);
                    string startAfterId = string.Empty;
                    //string startParentnewVer = Services.ApplicationServices.VersioningSvc.GetNextVersion(startParentId, "revise");
                    //string resul2t = Services.ApplicationServices.ActionSvc.ActionValidate("revise", startParentId, UserInfo.UserID);
                    //MessageBox.Show("가능 : " + resul2t);
                    string realIdcheckstate = Convert.ToString(startdr["pstate"]).ToLower();
                    string ecoAction = Convert.ToString(startdr["S$ECOACTION"]);
                    string ECOQuentity = Convert.ToString(startdr["S$ECOQuentity"]);
                    string ECOLOCATION = Convert.ToString(startdr["S$ECOLOCATION"]);
                    string ECOALTERNATE = Convert.ToString(startdr["S$ECOALTERNATE"]);
                    string pNumber = Convert.ToString(startdr["P$number"]);


                    string number = string.Empty;
                    string modelname = string.Empty;
                    string version = string.Empty;
                    Adaptive.Windows.Waiting.Show();




                    if (ecoAction == "ADD" || ecoAction == "DEL")
                    {
                        startAfterId = startFromId;
                    }
                    else
                    {
                        startAfterId = Services.ApplicationServices.DataSvc.ExecuteScalar("select urid from part_info where P$Number = '" + Convert.ToString(startdr["S$ECOAFTERNUM"]) + "' AND version = '" + Convert.ToString(startdr["S$ECOAFTERVER"]) + "'");

                    }

                    //MessageBox.Show(realIdcheckstate);

                    if (realIdcheckstate == "freezing")
                    {

                    }
                    else
                    {
                        if (realIdcheckstate == "underrevision" || realIdcheckstate == "checkedout")
                        {

                            startParentnewId = Convert.ToString(startdr["parent_id"]);



                            if (Services.ApplicationServices.StructureSvc.IsExistStructure(startParentnewId, "EBOM", -1, startAfterId) == true)
                            {
                                //이미 연결된 정보 팝업..
                            }
                            else
                            {

                                int sequence = Services.ApplicationServices.StructureSvc.GetNewSequence(startParentnewId, 10, 10);

                                //MessageBox.Show("1 : " + Services.ApplicationServices.ObjectSvc.GetPropertyValue(startParentnewId, "number") + ":" + Services.ApplicationServices.ObjectSvc.GetPropertyValue(startAfterId, "number"));
                                if (ecoAction == "ADD")
                                {
                                    string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, sequence, startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                    //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);
                                }


                                if (ecoAction == "DEL")
                                {
                                    //해당 어셈블리에 삭제만
                                    Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startAfterId);
                                }



                                if (ecoAction == "ALT")
                                {
                                    string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, sequence, startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                    //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);

                                    // 해당 어셈블리에 추가만
                                }



                                if (ecoAction == "CHG")
                                {
                                    // 해당 어셈블리에섭 부품 대체
                                    string chgsequence = Services.ApplicationServices.DataSvc.ExecuteScalar("select sequence  from structure_info where parent_id = '" + startParentId + "' AND child_id = '" + startFromId + "'");

                                    string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, Convert.ToInt32(chgsequence), startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                    Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startFromId);
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                    Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                    //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);

                                }




                            }
                        }
                        else
                        {
                            //MessageBox.Show(startParentnewVer + " : " + startParentnewId);

                            string startParentSate = Services.ApplicationServices.DataSvc.ExecuteScalar("select state from part_info where P$number = '" + pNumber + "' and latest = '1' order by version_sequence desc");
                            string startParentLatestId = Services.ApplicationServices.DataSvc.ExecuteScalar("select urid from part_info where P$number = '" + pNumber + "' and latest = '1' order by version_sequence desc");

                            //MessageBox.Show(startParentSate + " : " + startParentLatestId);
                            startParentnewId = Convert.ToString(startdr["parent_id"]);
                            if (startParentSate == "freezing")
                            {

                            }
                            else
                            {
                                if (startParentSate == "released")
                                {
                                    if (startParentnewId == startParentLatestId)
                                    {

                                        //string versionId = new URID(startParentnewId).

                                        //Adaptive.Search.SearchBuilder sb = new  Adaptive.Search.SearchBuilder("PART","PART",UserInfo.UserID);
                                        //sb.AppendProperty("latest", "1");
                                        //sb.AppendProperty("P$number", ""); //검색하여 가져오는 API Append 하나에 검색 조건 추가
                                        //DataTable tb =  Services.ApplicationServices.SearchSvc.GetSearchList(sb.ToXml());

                                        string result = Services.ApplicationServices.ActionSvc.ActionExecute("revise", "Application", startParentLatestId, UserInfo.UserID, "");
                                        ActionResult actionResult = Adaptive.Windows.ActionResultCheck.Check(result);
                                        startParentnewId = actionResult.ResultValue;
                                        string resultCheckedn = Services.ApplicationServices.ActionSvc.ActionExecute("CheckIn", "Application", startParentnewId, UserInfo.UserID, "");




                                        if (Services.ApplicationServices.StructureSvc.IsExistStructure(startParentnewId, "EBOM", -1, startAfterId) == true)
                                        {
                                            //이미 연결된 정보 팝업..
                                        }
                                        else
                                        {
                                            int sequence = Services.ApplicationServices.StructureSvc.GetNewSequence(startParentId, 10, 10);

                                            Services.ApplicationServices.StructureSvc.CopyStructure(startParentId, startParentnewId, UserInfo.UserID);

                                            if (ecoAction == "ADD")
                                            {
                                                string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, sequence, startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                                //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);

                                            }


                                            if (ecoAction == "DEL")
                                            {
                                                //해당 어셈블리에 삭제만
                                                Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startAfterId);
                                            }



                                            if (ecoAction == "ALT")
                                            {
                                                string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, sequence, startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                                //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);

                                                // 해당 어셈블리에 추가만
                                            }



                                            if (ecoAction == "CHG")
                                            {
                                                // 해당 어셈블리에섭 부품 대체

                                                string chgsequence = Services.ApplicationServices.DataSvc.ExecuteScalar("select sequence  from structure_info where parent_id = '" + startParentId + "' AND child_id = '" + startFromId + "'");
                                                //MessageBox.Show("1");
                                                string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, Convert.ToInt32(chgsequence), startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                                Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startFromId);
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                                Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                                //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);

                                            }

                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("승인된 최신 버전의 어셈블리가 아닙니다.");
                                    }

                                }
                                else
                                {
                                    startParentnewId = startParentLatestId;
                                    //startParentnewId = Convert.ToString(startdr["parent_id"]);
                                    if (Services.ApplicationServices.StructureSvc.IsExistStructure(startParentLatestId, "EBOM", -1, startAfterId) == true)
                                    {
                                        //이미 연결된 정보 팝업..
                                    }
                                    else
                                    {
                                        int sequence = Services.ApplicationServices.StructureSvc.GetNewSequence(startParentId, 10, 10);


                                        if (ecoAction == "ADD")
                                        {
                                            string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, sequence, startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                            //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);

                                        }


                                        if (ecoAction == "DEL")
                                        {
                                            //해당 어셈블리에 삭제만
                                            Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startAfterId);
                                        }



                                        if (ecoAction == "ALT")
                                        {

                                            string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, sequence, startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                            //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);

                                            // 해당 어셈블리에 추가만
                                        }



                                        if (ecoAction == "CHG")
                                        {
                                            // 해당 어셈블리에섭 부품 대체
                                            string chgsequence = Services.ApplicationServices.DataSvc.ExecuteScalar("select sequence  from structure_info where parent_id = '" + startParentId + "' AND child_id = '" + startFromId + "'");

                                            string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, Convert.ToInt32(chgsequence), startAfterId, Convert.ToInt64(startdr["S$ECOQuentity"]));
                                            Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startFromId);
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "Quantity", ECOQuentity);
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$Location", ECOLOCATION);
                                            Services.ApplicationServices.ObjectSvc.UpdatePropertyValue(partStructureId, "S$AlternativeOrder", ECOALTERNATE);
                                            //Services.ApplicationServices.RelationSvc.CreateRelation("UsedChangeManage", UserInfo.UserID, startAfterId, this.topId);

                                        }
                                    }

                                }
                            }
                        }

                        // 개정 시작


                        if (structureInfoDtStu2.Rows.Count != 0)
                        {
                            BOMecostart1(startParentId, startParentnewId);

                        }
                        else
                        {
                            //string relCheckst = string.Format("select*from relation_info where to_id = '" + this.TopID + "' AND from_id = '" + startParentnewId + "'");

                            //     DataTable relCheck = Services.ApplicationServices.DataSvc.ExecuteDataTable(relCheckst);
                            //     if (relCheck.Rows.Count != 0)
                            //     {
                            //         Services.ApplicationServices.RelationSvc.CreateRelation("AttachedPart", UserInfo.UserID, startParentnewId, this.TopID);
                            //         //최상위 이기 때문에 릴레이션 추가.
                            //     }
                        }


                        //2015-05-26 일괄개정 후 개정된 최상위 어셈블리 연결 로직
                        //DataTable dt = Services.ApplicationServices.StructureSvc.GetTops(startParentnewId, "EBOM", "Part");

                        //foreach (DataRow dr in dt.Rows)
                        //{
                        //    //MessageBox.Show("!11");
                        //    //DataTable dtparent = Services.ApplicationServices.DataSvc.ExecuteDataTable("select* FROM structure_info s INNER JOIN part_info$ c on s.child_id  = c.part_id  WHERE s.child_id =  '" + dr["parent_id"].ToString() + "' AND s.structure_id LIKE 'S$EBOM%' AND p$latest = 1 ");
                        //    ////MessageBox.Show(dr["p$Latest"].ToString());
                        //    //if (dtparent.Rows.Count == 0)
                        //    //{
                        //        //Application.DoEvents();
                        //        //MessageBox.Show(dr["p$Latest"].ToString());
                        //        Application.DoEvents();
                        //        if (dr["state"].ToString() != "freezing")
                        //        {
                        //            if (dr["Latest"].ToString() == "True")
                        //            {

                        //                //number = dr["p$number"].ToString();
                        //                //modelname = dr["p$modelname"].ToString();
                        //                //version = dr["version"].ToString();
                        //                DataTable dtrel = Services.ApplicationServices.DataSvc.ExecuteDataTable("select* FROM relation_entity where from_id = '" + dr["urid"].ToString() + "' AND to_id = '" + topid + "'");
                        //                if (dtrel.Rows.Count == 0)
                        //                {
                        //                    Services.ApplicationServices.RelationSvc.CreateRelation("AttachedPart", UserInfo.UserID, dr["urid"].ToString(), topid);
                        //                }
                        //                //MessageBox.Show("1");
                        //            }
                        //        }
                        //    //}

                        //}
                    }

                }
                catch (Exception ex)
                {
                    
                    Adaptive.DebugConsole.WriteLine(ex.StackTrace + ", " + ex.Message);
                  
                }



            }

            Adaptive.Windows.Waiting.Close();
            if (checkcount == "2")
            {
                MessageBox.Show("상위 어셈블리에 문제가 있습니다.");
            }
            else
            {
                MessageBox.Show("적용 되었습니다. 결재 받으십시오");
            }
        }

        public void BOMecostart1(string startFromId, string startToId)
        {
            //해당 어셈블리와 역전개시 최신 구조만 리비전
            string sqlStu = string.Format("select*, p.state as pstate from structure_info as s INNER JOIN  part_info as p on s.parent_id = p.urid where s.child_id = '" + startFromId + "' AND p.latest = '1'");
            DataTable structureInfoDtStu = Services.ApplicationServices.DataSvc.ExecuteDataTable(sqlStu);
            //MessageBox.Show("1");
            //MessageBox.Show("2");

            foreach (DataRow startParentdr in structureInfoDtStu.Rows)
            {
                Application.DoEvents();
                string startParentnewId = string.Empty;
                string startParentId = Convert.ToString(startParentdr["parent_id"]);
                string sqlStu2 = string.Format("select*, p.state as pstate from structure_info as s INNER JOIN  part_info as p on s.parent_id = p.urid where s.child_id = '" + startParentId + "' AND p.latest = '1'");
                DataTable structureInfoDtStu2 = Services.ApplicationServices.DataSvc.ExecuteDataTable(sqlStu2);
                //MessageBox.Show("2");
                string startAfterId = startToId;
                string startfromId2 = startFromId;
                string realIdcheckstate = Convert.ToString(startParentdr["pstate"]).ToLower();
                string chgsequence = Services.ApplicationServices.DataSvc.ExecuteScalar("select sequence  from structure_info where parent_id = '" + startParentId + "' AND child_id = '" + startFromId + "'");


                //MessageBox.Show(realIdcheckstate);
                if (realIdcheckstate == "freezing")
                {

                }
                else
                {
                    if (realIdcheckstate == "underrevision" || realIdcheckstate == "checkedout")
                    {
                        startParentnewId = Convert.ToString(startParentdr["parent_id"]);

                        if (Services.ApplicationServices.StructureSvc.IsExistStructure(startParentnewId, "EBOM", -1, startAfterId) == true)
                        {
                            //이미 연결된 정보 팝업..
                        }
                        else
                        {
                            int sequence = Services.ApplicationServices.StructureSvc.GetNewSequence(startParentnewId, 10, 10);

                            //MessageBox.Show("1 : " + Services.ApplicationServices.ObjectSvc.GetPropertyValue(startParentnewId, "number") + ":" + Services.ApplicationServices.ObjectSvc.GetPropertyValue(startAfterId, "number"));

                            string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, Convert.ToInt32(chgsequence), startAfterId, Convert.ToInt64(startParentdr["Quantity"]));

                            Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startfromId2);
                        }

                        if (structureInfoDtStu2.Rows.Count != 0)
                        {
                            //BOMecostart2(startfromId2, startParentnewId);
                        }
                        else
                        {
                            //MessageBox.Show("2");
                            string relCheckst = string.Format("select*from relation_info where to_id = '" + topid + "' AND from_id = '" + startParentnewId + "'");

                            DataTable relCheck = Services.ApplicationServices.DataSvc.ExecuteDataTable(relCheckst);
                            if (relCheck.Rows.Count == 0)
                            {
                                Services.ApplicationServices.RelationSvc.CreateRelation("AttachedPart", UserInfo.UserID, startParentnewId, topid);
                                //최상위 이기 때문에 릴레이션 추가.
                            }
                        }


                    }
                    else
                    {
                        string pNumber = Convert.ToString(startParentdr["P$number"]);

                        string startParentSate = Services.ApplicationServices.DataSvc.ExecuteScalar("select state from part_info where P$number = '" + pNumber + "' and latest = '1' order by version_sequence desc");
                        string startParentLatestId = Services.ApplicationServices.DataSvc.ExecuteScalar("select urid from part_info where P$number = '" + pNumber + "' and latest = '1' order by version_sequence desc");

                        if (startParentSate == "freezing")
                        {

                        }
                        else
                        {
                            if (startParentSate == "released")
                            {

                                startParentnewId = startParentLatestId;
                                startAfterId = startToId;
                                startfromId2 = startFromId;
                                //MessageBox.Show(startAfterId);              
                                string startParentnewVer = Services.ApplicationServices.VersioningSvc.GetNextVersion(startParentId, "revise");

                                //string resul2t = Services.ApplicationServices.ActionSvc.ActionValidate("revise",startParentId, UserInfo.UserID);

                                string result2 = Services.ApplicationServices.ActionSvc.ActionValidate("revise", startParentId, UserInfo.UserID);
                                string result = Services.ApplicationServices.ActionSvc.ActionExecute("revise", "Application", startParentId, UserInfo.UserID, "");
                                ActionResult actionResult = Adaptive.Windows.ActionResultCheck.Check(result);
                                startParentnewId = actionResult.ResultValue;
                                string resultCheckedn = Services.ApplicationServices.ActionSvc.ActionExecute("CheckIn", "Application", startParentnewId, UserInfo.UserID, "");

                                //MessageBox.Show(startParentnewVer + " : " + startParentnewId);
                                if (Services.ApplicationServices.StructureSvc.IsExistStructure(startParentnewId, "EBOM", -1, startAfterId) == true)
                                {
                                    //이미 연결된 정보 팝업..
                                }
                                else
                                {
                                    int sequence = Services.ApplicationServices.StructureSvc.GetNewSequence(startParentId, 10, 10);

                                    Services.ApplicationServices.StructureSvc.CopyStructure(startParentId, startParentnewId, UserInfo.UserID);
                                    //MessageBox.Show("2 : " + Services.ApplicationServices.ObjectSvc.GetPropertyValue(startParentnewId, "number") + ":" + Services.ApplicationServices.ObjectSvc.GetPropertyValue(startAfterId, "number"));
                                    string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, Convert.ToInt32(chgsequence), startAfterId, Convert.ToInt64(startParentdr["Quantity"]));

                                    Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startfromId2);

                                }

                                if (structureInfoDtStu2.Rows.Count != 0)
                                {
                                    BOMecostart2(startParentId, startParentnewId);
                                }
                                else
                                {
                                    //   string relCheckst = string.Format("select*from relation_info where to_id = '" + this.TopID + "' AND from_id = '" + startParentnewId + "'");
                                    //   //MessageBox.Show("2");
                                    //DataTable relCheck = Services.ApplicationServices.DataSvc.ExecuteDataTable(relCheckst);
                                    //if (relCheck.Rows.Count == 0)
                                    //{
                                    //    Services.ApplicationServices.RelationSvc.CreateRelation("AttachedPart", UserInfo.UserID, startParentnewId, this.TopID);
                                    //    //최상위 이기 때문에 릴레이션 추가.
                                    //}
                                }


                            }
                            else
                            {
                                startParentnewId = Convert.ToString(startParentdr["parent_id"]);

                                if (Services.ApplicationServices.StructureSvc.IsExistStructure(startParentnewId, "EBOM", -1, startAfterId) == true)
                                {
                                    //이미 연결된 정보 팝업..
                                }
                                else
                                {

                                    int sequence = Services.ApplicationServices.StructureSvc.GetNewSequence(startParentnewId, 10, 10);


                                    string partStructureId = Services.ApplicationServices.StructureSvc.CreateStructure(startParentnewId, "EBOM", UserInfo.UserID, Convert.ToInt32(chgsequence), startAfterId, Convert.ToInt64(startParentdr["Quantity"]));

                                    Services.ApplicationServices.StructureSvc.DeleteStructure(startParentnewId, startfromId2);
                                }

                                if (structureInfoDtStu2.Rows.Count != 0)
                                {
                                    //BOMecostart2(startParentId, startParentnewId);
                                }
                                else
                                {
                                    //MessageBox.Show("2");
                                    //   string relCheckst = string.Format("select*from relation_info where to_id = '" + this.TopID + "' AND from_id = '" + startParentnewId + "'");

                                    //DataTable relCheck = Services.ApplicationServices.DataSvc.ExecuteDataTable(relCheckst);
                                    //if (relCheck.Rows.Count == 0)
                                    //{
                                    //    Services.ApplicationServices.RelationSvc.CreateRelation("AttachedPart", UserInfo.UserID, startParentnewId, this.TopID);
                                    //    //최상위 이기 때문에 릴레이션 추가.
                                    //}
                                }



                            }
                        }


                    }
                }


            }




        }


        public void BOMecostart2(string fromid, string toid)
        {
            BOMecostart1(fromid, toid);
            //BOMecostart2(string startParentId);
            //BOMecostart1(startParentId, startParentnewId);;

            //MessageBox.Show("2");
        }



    }
}
