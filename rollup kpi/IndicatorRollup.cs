using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Security;

namespace rollup_kpi
{
    class IndicatorRollup
    {
        private string userName = "admin.sharepoint@mca-indonesia.go.id";
        private string password = "admin123$";

        private string checkNull(object item)
        {
            try
            {
                return item.ToString();
            }
            catch
            {
                return "";
            }
        }

        private double checkNullDouble(object item)
        {
            try
            {
                return Convert.ToDouble(item.ToString());
            }
            catch
            {
                return 0;
            }
        }

        private int checkNullInt(object item)
        {
            try
            {
                return Convert.ToInt32(item.ToString());
            }
            catch
            {
                return 0;
            }
        }

        private DateTime checkNullDateTimeP(object item)
        {
            return Convert.ToDateTime(item.ToString());
        }

        private DateTime? checkNullDateTime(object item)
        {
            try
            {
                return new DateTime?(Convert.ToDateTime(item.ToString()));
            }
            catch
            {
                return null;
            }
        }

        void getSubIndicator(string CurUrl, string indicator, ref double[] target, ref double[] Qach, ref double[] Ach)
        {
            using (ClientContext context = new ClientContext(CurUrl))
            {
                SecureString secureString = new SecureString();
                password.ToList().ForEach(secureString.AppendChar);
                context.Credentials = new SharePointOnlineCredentials(userName, secureString);
                Site site = context.Site;
                context.Load(site);
                context.ExecuteQuery();

                List byTitle = context.Web.Lists.GetByTitle("Sub Indicator");
                CamlQuery query = new CamlQuery();
                query.ViewXml = @"
                <View>
                    <Query>
                       <Where>
                          <Eq>
                             <FieldRef Name='Indicator' />
                             <Value Type='Lookup'>" + indicator + @"</Value>
                          </Eq>
                       </Where>
                    </Query>
                </View>";
                Microsoft.SharePoint.Client.ListItemCollection clientObject = byTitle.GetItems(query);
                context.Load<Microsoft.SharePoint.Client.ListItemCollection>(clientObject);
                context.ExecuteQuery();

                for (int i = 0; i < 21; i++)
                {
                    target[i] = 0;
                }


                for (int i = 0; i < 66; i++)
                {
                    Ach[i] = 0;
                }

                int countRow = clientObject.Count;

                foreach (Microsoft.SharePoint.Client.ListItem item in clientObject)
                {
                    target[0] += checkNullDouble(item["Q01Target"]);
                    target[1] += checkNullDouble(item["Q02Target"]);
                    target[2] += checkNullDouble(item["Q03Target"]);
                    target[3] += checkNullDouble(item["Q04Target"]);
                    target[4] += checkNullDouble(item["Q05Target"]);
                    target[5] += checkNullDouble(item["Q06Target"]);
                    target[6] += checkNullDouble(item["Q07Target"]);
                    target[7] += checkNullDouble(item["Q08Target"]);
                    target[8] += checkNullDouble(item["Q09Target"]);
                    target[9] += checkNullDouble(item["Q10Target"]);
                    target[10] += checkNullDouble(item["Q11Target"]);
                    target[11] += checkNullDouble(item["Q12Target"]);
                    target[12] += checkNullDouble(item["Q13Target"]);
                    target[13] += checkNullDouble(item["Q14Target"]);
                    target[14] += checkNullDouble(item["Q15Target"]);
                    target[15] += checkNullDouble(item["Q16Target"]);
                    target[16] += checkNullDouble(item["Q17Target"]);
                    target[17] += checkNullDouble(item["Q18Target"]);
                    target[18] += checkNullDouble(item["Q19Target"]);
                    target[19] += checkNullDouble(item["Q20Target"]);
                    target[20] += checkNullDouble(item["Q21Target"]);

                    Qach[0] += checkNullDouble(item["Q01Achievement"]);
                    Qach[1] += checkNullDouble(item["Q02Achievement"]);
                    Qach[2] += checkNullDouble(item["Q03Achievement"]);
                    Qach[3] += checkNullDouble(item["Q04Achievement"]);
                    Qach[4] += checkNullDouble(item["Q05Achievement"]);
                    Qach[5] += checkNullDouble(item["Q06Achievement"]);
                    Qach[6] += checkNullDouble(item["Q07Achievement"]);
                    Qach[7] += checkNullDouble(item["Q08Achievement"]);
                    Qach[8] += checkNullDouble(item["Q09Achievement"]);
                    Qach[9] += checkNullDouble(item["Q10Achievement"]);
                    Qach[10] += checkNullDouble(item["Q11Achievement"]);
                    Qach[11] += checkNullDouble(item["Q12Achievement"]);
                    Qach[12] += checkNullDouble(item["Q13Achievement"]);
                    Qach[13] += checkNullDouble(item["Q14Achievement"]);
                    Qach[14] += checkNullDouble(item["Q15Achievement"]);
                    Qach[15] += checkNullDouble(item["Q16Achievement"]);
                    Qach[16] += checkNullDouble(item["Q17Achievement"]);
                    Qach[17] += checkNullDouble(item["Q18Achievement"]);
                    Qach[18] += checkNullDouble(item["Q19Achievement"]);
                    Qach[19] += checkNullDouble(item["Q20Achievement"]);
                    Qach[20] += checkNullDouble(item["Q21Achievement"]);

                    Ach[0] += checkNullDouble(item["_x0041_pr13"]);
                    Ach[1] += checkNullDouble(item["_x004d_ay13"]);
                    Ach[2] += checkNullDouble(item["_x004a_un13"]);
                    Ach[3] += checkNullDouble(item["_x004a_ul13"]);
                    Ach[4] += checkNullDouble(item["_x0041_ug13"]);
                    Ach[5] += checkNullDouble(item["_x0053_ep13"]);
                    Ach[6] += checkNullDouble(item["_x004f_ct13"]);
                    Ach[7] += checkNullDouble(item["_x004e_ov13"]);
                    Ach[8] += checkNullDouble(item["_x0044_ec13"]);
                    Ach[9] += checkNullDouble(item["_x004a_an14"]);
                    Ach[10] += checkNullDouble(item["_x0046_eb14"]);
                    Ach[11] += checkNullDouble(item["_x004d_ar14"]);
                    Ach[12] += checkNullDouble(item["_x0041_pr14"]);
                    Ach[13] += checkNullDouble(item["_x004d_ay14"]);
                    Ach[14] += checkNullDouble(item["_x004a_un14"]);
                    Ach[15] += checkNullDouble(item["_x004a_ul14"]);
                    Ach[16] += checkNullDouble(item["_x0041_ug14"]);
                    Ach[17] += checkNullDouble(item["_x0053_ep14"]);
                    Ach[18] += checkNullDouble(item["_x004f_ct14"]);
                    Ach[19] += checkNullDouble(item["_x004e_ov14"]);
                    Ach[20] += checkNullDouble(item["_x0044_ec14"]);
                    Ach[21] += checkNullDouble(item["_x004a_an15"]);
                    Ach[22] += checkNullDouble(item["_x0046_eb15"]);
                    Ach[23] += checkNullDouble(item["_x004d_ar15"]);
                    Ach[24] += checkNullDouble(item["_x0041_pr15"]);
                    Ach[25] += checkNullDouble(item["_x004d_ay15"]);
                    Ach[26] += checkNullDouble(item["_x004a_un15"]);
                    Ach[27] += checkNullDouble(item["_x004a_ul15"]);
                    Ach[28] += checkNullDouble(item["_x0041_ug15"]);
                    Ach[29] += checkNullDouble(item["_x0053_ep15"]);
                    Ach[30] += checkNullDouble(item["_x004f_ct15"]);
                    Ach[31] += checkNullDouble(item["_x004e_ov15"]);
                    Ach[32] += checkNullDouble(item["_x0044_ec15"]);
                    Ach[33] += checkNullDouble(item["_x004a_an16"]);
                    Ach[34] += checkNullDouble(item["_x0046_eb16"]);
                    Ach[35] += checkNullDouble(item["_x004d_ar16"]);
                    Ach[36] += checkNullDouble(item["_x0041_pr16"]);
                    Ach[37] += checkNullDouble(item["_x004d_ay16"]);
                    Ach[38] += checkNullDouble(item["_x004a_un16"]);
                    Ach[39] += checkNullDouble(item["_x004a_ul16"]);
                    Ach[40] += checkNullDouble(item["_x0041_ug16"]);
                    Ach[41] += checkNullDouble(item["_x0053_ep16"]);
                    Ach[42] += checkNullDouble(item["_x004f_ct16"]);
                    Ach[43] += checkNullDouble(item["_x004e_ov16"]);
                    Ach[44] += checkNullDouble(item["_x0044_ec16"]);
                    Ach[45] += checkNullDouble(item["_x004a_an17"]);
                    Ach[46] += checkNullDouble(item["_x0046_eb17"]);
                    Ach[47] += checkNullDouble(item["_x004d_ar17"]);
                    Ach[48] += checkNullDouble(item["_x0041_pr17"]);
                    Ach[49] += checkNullDouble(item["_x004d_ay17"]);
                    Ach[50] += checkNullDouble(item["_x004a_un17"]);
                    Ach[51] += checkNullDouble(item["_x004a_ul17"]);
                    Ach[52] += checkNullDouble(item["_x0041_ug17"]);
                    Ach[53] += checkNullDouble(item["_x0053_ep17"]);
                    Ach[54] += checkNullDouble(item["_x004f_ct17"]);
                    Ach[55] += checkNullDouble(item["_x004e_ov17"]);
                    Ach[56] += checkNullDouble(item["_x0044_ec17"]);
                    Ach[57] += checkNullDouble(item["_x004a_an18"]);
                    Ach[58] += checkNullDouble(item["_x0046_eb18"]);
                    Ach[59] += checkNullDouble(item["_x004d_ar18"]);
                    Ach[60] += checkNullDouble(item["_x0041_pr18"]);
                    Ach[61] += checkNullDouble(item["_x004d_ay18"]);
                    Ach[62] += checkNullDouble(item["_x004a_un18"]);
                    Ach[63] += checkNullDouble(item["_x004a_ul18"]);
                    Ach[64] += checkNullDouble(item["_x0041_ug18"]);
                    Ach[65] += checkNullDouble(item["_x0053_ep18"]);

                }
            }
        }

        public void GetIndicatorWithRollup(string CurUrl)
        {
            using (ClientContext context = new ClientContext(CurUrl))
            {
                SecureString secureString = new SecureString();
                password.ToList().ForEach(secureString.AppendChar);
                context.Credentials = new SharePointOnlineCredentials(userName, secureString);
                Site site = context.Site;
                context.Load(site);
                context.ExecuteQuery();

                List byTitle = context.Web.Lists.GetByTitle("Indicator");
                CamlQuery query = new CamlQuery();
                query.ViewXml = @"
                <View>
                    <Query>
                       <Where>
                          <Eq>
                             <FieldRef Name='Rollup' />
                             <Value Type='Boolean'>1</Value>
                          </Eq>
                       </Where>
                    </Query>
                </View>";
                Microsoft.SharePoint.Client.ListItemCollection clientObject = byTitle.GetItems(query);
                context.Load<Microsoft.SharePoint.Client.ListItemCollection>(clientObject);
                context.ExecuteQuery();

                double[] target = new double[21];
                double[] ach = new double[66];
                double[] qach = new double[21];

                

                foreach (Microsoft.SharePoint.Client.ListItem item in clientObject)
                {
                    for (int i = 0; i < 21; i++)
                    {
                        target[i] = 0;
                        qach[i] = 0;
                    }


                    for (int i = 0; i < 66; i++)
                    {
                        ach[i] = 0;
                    }

                    getSubIndicator(CurUrl, checkNull(item["Title"]), ref target, ref qach, ref ach);

                    //item["Q01Target"] = target[0];
                    //item["Q02Target"] = target[1];
                    //item["Q03Target"] = target[2];
                    //item["Q04Target"] = target[3];
                    //item["Q05Target"] = target[4];
                    //item["Q06Target"] = target[5];
                    //item["Q07Target"] = target[6];
                    //item["Q08Target"] = target[7];
                    //item["Q09Target"] = target[8];
                    //item["Q10Target"] = target[9];
                    //item["Q11Target"] = target[10];
                    //item["Q12Target"] = target[11];
                    //item["Q13Target"] = target[12];
                    //item["Q14Target"] = target[13];
                    //item["Q15Target"] = target[14];
                    //item["Q16Target"] = target[15];
                    //item["Q17Target"] = target[16];
                    //item["Q18Target"] = target[17];
                    //item["Q19Target"] = target[18];
                    //item["Q20Target"] = target[19];
                    //item["Q21Target"] = target[20];

                    item["Q01Achievement"] = qach[0];
                    item["Q02Achievement"] = qach[1];
                    item["Q03Achievement"] = qach[2];
                    item["Q04Achievement"] = qach[3];
                    item["Q05Achievement"] = qach[4];
                    item["Q06Achievement"] = qach[5];
                    item["Q07Achievement"] = qach[6];
                    item["Q08Achievement"] = qach[7];
                    item["Q09Achievement"] = qach[8];
                    item["Q10Achievement"] = qach[9];
                    item["Q11Achievement"] = qach[10];
                    item["Q12Achievement"] = qach[11];
                    item["Q13Achievement"] = qach[12];
                    item["Q14Achievement"] = qach[13];
                    item["Q15Achievement"] = qach[14];
                    item["Q16Achievement"] = qach[15];
                    item["Q17Achievement"] = qach[16];
                    item["Q18Achievement"] = qach[17];
                    item["Q19Achievement"] = qach[18];
                    item["Q20Achievement"] = qach[19];
                    item["Q21Achievement"] = qach[20];

                    item["_x0041_pr13"] = ach[0];
                    item["_x004d_ay13"] = ach[1];
                    item["_x004a_un13"] = ach[2];
                    item["_x004a_ul13"] = ach[3];
                    item["_x0041_ug13"] = ach[4];
                    item["_x0053_ep13"] = ach[5];
                    item["_x004f_ct13"] = ach[6];
                    item["_x004e_ov13"] = ach[7];
                    item["_x0044_ec13"] = ach[8];
                    item["_x004a_an14"] = ach[9];
                    item["_x0046_eb14"] = ach[10];
                    item["_x004d_ar14"] = ach[11];
                    item["_x0041_pr14"] = ach[12];
                    item["_x004d_ay14"] = ach[13];
                    item["_x004a_un14"] = ach[14];
                    item["_x004a_ul14"] = ach[15];
                    item["_x0041_ug14"] = ach[16];
                    item["_x0053_ep14"] = ach[17];
                    item["_x004f_ct14"] = ach[18];
                    item["_x004e_ov14"] = ach[19];
                    item["_x0044_ec14"] = ach[20];
                    item["_x004a_an15"] = ach[21];
                    item["_x0046_eb15"] = ach[22];
                    item["_x004d_ar15"] = ach[23];
                    item["_x0041_pr15"] = ach[24];
                    item["_x004d_ay15"] = ach[25];
                    item["_x004a_un15"] = ach[26];
                    item["_x004a_ul15"] = ach[27];
                    item["_x0041_ug15"] = ach[28];
                    item["_x0053_ep15"] = ach[29];
                    item["_x004f_ct15"] = ach[30];
                    item["_x004e_ov15"] = ach[31];
                    item["_x0044_ec15"] = ach[32];
                    item["_x004a_an16"] = ach[33];
                    item["_x0046_eb16"] = ach[34];
                    item["_x004d_ar16"] = ach[35];
                    item["_x0041_pr16"] = ach[36];
                    item["_x004d_ay16"] = ach[37];
                    item["_x004a_un16"] = ach[38];
                    item["_x004a_ul16"] = ach[39];
                    item["_x0041_ug16"] = ach[40];
                    item["_x0053_ep16"] = ach[41];
                    item["_x004f_ct16"] = ach[42];
                    item["_x004e_ov16"] = ach[43];
                    item["_x0044_ec16"] = ach[44];
                    item["_x004a_an17"] = ach[45];
                    item["_x0046_eb17"] = ach[46];
                    item["_x004d_ar17"] = ach[47];
                    item["_x0041_pr17"] = ach[48];
                    item["_x004d_ay17"] = ach[49];
                    item["_x004a_un17"] = ach[50];
                    item["_x004a_ul17"] = ach[51];
                    item["_x0041_ug17"] = ach[52];
                    item["_x0053_ep17"] = ach[53];
                    item["_x004f_ct17"] = ach[54];
                    item["_x004e_ov17"] = ach[55];
                    item["_x0044_ec17"] = ach[56];
                    item["_x004a_an18"] = ach[57];
                    item["_x0046_eb18"] = ach[58];
                    item["_x004d_ar18"] = ach[59];
                    item["_x0041_pr18"] = ach[60];
                    item["_x004d_ay18"] = ach[61];
                    item["_x004a_un18"] = ach[62];
                    item["_x004a_ul18"] = ach[63];
                    item["_x0041_ug18"] = ach[64];
                    item["_x0053_ep18"] = ach[65];

                    item.Update();
                    context.ExecuteQuery();
                }

            }
        }
    }
}
