using DocumentFormat.OpenXml.Spreadsheet;
using MailKit.Security;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using MimeKit;
using Org.BouncyCastle.Asn1.Ocsp;
using Asset_Management.Function;
using Asset_Management.Models;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Net.Mail;
using System.Security.Claims;
using SmtpClient = MailKit.Net.Smtp.SmtpClient;

namespace Asset_Management.Controllers
{
    [Authorize(Policy = "RequireAny")]
    public class EmailController : Controller
    {
        private string email_name = "ASSET MANAGEMENT";
        private string email_account = "YOUR-EMAIL-ACCOUNT";
        private string email_token = "YOUR-EMAIL-TOKEN";
        private string DbConnection()
        {
            var dbAccess = new DatabaseAccessLayer();
            string dbString = dbAccess.ConnectionString;
            return dbString;
        }
        private string GetLogoHtml()
        {
            return $@"<div style='display: flex; justify-content: space-between; align-items: center; background-image: url(https://i.imgur.com/Om9ShPA.jpeg); background-size: cover; background-position: center; background-repeat: no-repeat; padding: 20px;'>
            <img src='https://i.imgur.com/P6RKZ5I.png' alt='Company Logo' style='height: 35px;'>
            <img src='https://i.imgur.com/AkZ47fM.png' alt='Web Logo' style='height: 35px;'>
            </div>";
        }

        public string EMAIL_REQ_GATEPASS(int id_gatepass, bool isDelegationInvolved = false, string originalApprover = "")
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string location = "";
                string type = "";
                string employee_name = "";
                string csr_to = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";
                string send_bcc = "leaantony1707@gmail.com";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, type, employee_name, csr_to, create_date, return_date, created_by, vendor_name, requestor_name, location FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                location = reader["location"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                                type = reader["type"].ToString() ?? "";
                                employee_name = reader["employee_name"].ToString() ?? "";
                                csr_to = reader["csr_to"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_cc = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    if (isDelegationInvolved && !string.IsNullOrEmpty(originalApprover))
                    {
                        using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 approval_by FROM tbl_gatepass_approval WHERE id_gatepass=@id_gatepass AND approval_status='OPEN' ORDER BY approval_no", conn))
                        {
                            cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                            var nextApproverSesaId = cmd.ExecuteScalar()?.ToString() ?? "";

                            if (!string.IsNullOrEmpty(nextApproverSesaId))
                            {
                                using (SqlCommand checkDelegationCmd = new SqlCommand("SELECT delegated_to FROM mst_users WHERE sesa_id = @sesa_id", conn))
                                {
                                    checkDelegationCmd.Parameters.AddWithValue("@sesa_id", nextApproverSesaId);
                                    var delegatedTo = checkDelegationCmd.ExecuteScalar()?.ToString() ?? "";

                                    if (!string.IsNullOrEmpty(delegatedTo))
                                    {
                                        using (SqlCommand cmd2 = new SqlCommand("SELECT name, email FROM mst_users WHERE sesa_id = @original_approver OR sesa_id = @delegated_to", conn))
                                        {
                                            cmd2.Parameters.AddWithValue("@original_approver", nextApproverSesaId);
                                            cmd2.Parameters.AddWithValue("@delegated_to", delegatedTo);

                                            var names = new List<string>();
                                            var emails = new List<string>();

                                            using (SqlDataReader reader = cmd2.ExecuteReader())
                                            {
                                                while (reader.Read())
                                                {
                                                    names.Add(reader["name"].ToString() ?? "");
                                                    emails.Add(reader["email"].ToString() ?? "");
                                                }
                                            }

                                            if (emails.Count > 0)
                                            {
                                                send_to = string.Join(";", emails);
                                                send_to_name = string.Join(" & ", names);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        using (SqlCommand cmd2 = new SqlCommand("SELECT name, email FROM mst_users WHERE sesa_id = @sesa_id", conn))
                                        {
                                            cmd2.Parameters.AddWithValue("@sesa_id", nextApproverSesaId);
                                            using (SqlDataReader reader = cmd2.ExecuteReader())
                                            {
                                                if (reader.Read())
                                                {
                                                    send_to_name = reader["name"].ToString() ?? "";
                                                    send_to = reader["email"].ToString() ?? "";
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 approval_by FROM tbl_gatepass_approval WHERE id_gatepass=@id_gatepass AND approval_status='OPEN' ORDER BY approval_no", conn))
                        {
                            cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                            var nextApproverSesaId = cmd.ExecuteScalar()?.ToString() ?? "";

                            if (!string.IsNullOrEmpty(nextApproverSesaId))
                            {
                                using (SqlCommand checkDelegationCmd = new SqlCommand("SELECT delegated_to FROM mst_users WHERE sesa_id = @sesa_id", conn))
                                {
                                    checkDelegationCmd.Parameters.AddWithValue("@sesa_id", nextApproverSesaId);
                                    var delegatedTo = checkDelegationCmd.ExecuteScalar()?.ToString() ?? "";

                                    if (!string.IsNullOrEmpty(delegatedTo))
                                    {
                                        using (SqlCommand cmd2 = new SqlCommand("SELECT name, email FROM mst_users WHERE sesa_id = @original_approver OR sesa_id = @delegated_to", conn))
                                        {
                                            cmd2.Parameters.AddWithValue("@original_approver", nextApproverSesaId);
                                            cmd2.Parameters.AddWithValue("@delegated_to", delegatedTo);

                                            var names = new List<string>();
                                            var emails = new List<string>();

                                            using (SqlDataReader reader = cmd2.ExecuteReader())
                                            {
                                                while (reader.Read())
                                                {
                                                    names.Add(reader["name"].ToString() ?? "");
                                                    emails.Add(reader["email"].ToString() ?? "");
                                                }
                                            }

                                            if (emails.Count > 0)
                                            {
                                                send_to = string.Join(";", emails);
                                                send_to_name = string.Join(" & ", names);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        using (SqlCommand cmd2 = new SqlCommand("SELECT name, email FROM mst_users WHERE sesa_id = @sesa_id", conn))
                                        {
                                            cmd2.Parameters.AddWithValue("@sesa_id", nextApproverSesaId);
                                            using (SqlDataReader reader = cmd2.ExecuteReader())
                                            {
                                                if (reader.Read())
                                                {
                                                    send_to_name = reader["name"].ToString() ?? "";
                                                    send_to = reader["email"].ToString() ?? "";
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    conn.Close();
                }

                if (send_to == "")
                {
                    return "No Open Approval";
                }
                else
                {
                    string approvalLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                    string tableRows = "";
                    if (category == "TRANSFER PLANT TO PLANT")
                    {
                        tableRows = $@"
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666;'>New Location</td>
                                <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{location}</td>
                            </tr>";
                    }
                    else if (category == "TRANSFER INSIDE PLANT")
                    {
                        tableRows = $@"
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666;'>Return Date</td>
                                <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{return_date}</td>
                            </tr>";
                    }
                    else if (category == "FA SALES - WITHOUT DISMANTLING")
                    {
                        string additionalField = "";
                        if (type == "VENDOR")
                        {
                            additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Vendor Name</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td></tr>";
                        }
                        else if (type == "EMPLOYEE")
                        {
                            additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Employee Name</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{employee_name}</td></tr>";
                        }
                        else
                        {
                            additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Destination To</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{csr_to}</td></tr>";
                        }
                        tableRows = $@"
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category} ({type})</td>
                            </tr>
                            {additionalField}";
                    }
                    else
                    {
                        tableRows = $@"
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                            </tr>";
                    }

                    string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #007bff; padding: 20px; border-bottom: 2px solid #007bff;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Approval Request</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p><strong>{requestor_name}</strong> has requested your approval for Gate Pass <strong>{gatepass_no}</strong>. Please review the details below.</p>
                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        {tableRows}
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{approvalLink}' target='_blank' style='background-color: #007bff; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#007bff'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>Approve Gate Pass</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{approvalLink}' style='color: #007bff; text-decoration: none;'>{approvalLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                    var builder = new BodyBuilder();
                    builder.HtmlBody = emailContent;

                    email.Body = builder.ToMessageBody();

                    MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                    email.From.Add(from);

                    foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress to = new MailboxAddress("", email_to.Trim());
                        email.To.Add(to);
                    }

                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }

                    foreach (string email_bcc in send_bcc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress bcc = new MailboxAddress("", email_bcc.Trim());
                        email.Bcc.Add(bcc);
                    }

                    MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                    email.Bcc.Add(bcc1);

                    email.Subject = $"ASSET MANAGEMENT - Approval Request - {gatepass_no}";

                    using (var client = new SmtpClient())
                    {
                        client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                        client.Authenticate(email_account, email_token);
                        client.Send(email);
                        client.Disconnect(true);
                    }

                    return "success";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string EMAIL_REQ_GATEPASS_RETURN(int id_gatepass)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string location = "";
                string type = "";
                string employee_name = "";
                string csr_to = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";
                string send_bcc = "leaantony1707@gmail.com";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, type, employee_name, csr_to, create_date, return_date, created_by, vendor_name, requestor_name, location FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                location = reader["location"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                                type = reader["type"].ToString() ?? "";
                                employee_name = reader["employee_name"].ToString() ?? "";
                                csr_to = reader["csr_to"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_cc = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id IN (SELECT TOP 1 approval_by FROM tbl_gatepass_hod WHERE id_gatepass=@id_gatepass AND approval_status='OPEN')", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    conn.Close();
                }

                if (send_to == "")
                {
                    return "No Open Approval";
                }
                else
                {
                    string approvalLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                    string tableRows = "";
                    if (category == "REPAIR TO VENDOR")
                    {
                        tableRows = $@"
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                            </tr>";
                    }
                    else
                    {
                        tableRows = $@"
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                            </tr>
                            <tr>
                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                            </tr>
                            <tr style='background-color: #fcfcfc;'>
                                <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                            </tr>";
                    }

                    string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #007bff; padding: 20px; border-bottom: 2px solid #007bff;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Return Approval Request</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p><strong>{requestor_name}</strong> has requested your approval for the return of Gate Pass <strong>{gatepass_no}</strong>. Please review the details below.</p>
                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        {tableRows}
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{approvalLink}' target='_blank' style='background-color: #007bff; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#007bff'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>Approve Gate Pass Return</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{approvalLink}' style='color: #007bff; text-decoration: none;'>{approvalLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                    var builder = new BodyBuilder();
                    builder.HtmlBody = emailContent;

                    email.Body = builder.ToMessageBody();

                    MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                    email.From.Add(from);

                    foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress to = new MailboxAddress("", email_to.Trim());
                        email.To.Add(to);
                    }

                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }

                    foreach (string email_bcc in send_bcc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress bcc = new MailboxAddress("", email_bcc.Trim());
                        email.Bcc.Add(bcc);
                    }

                    MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                    email.Bcc.Add(bcc1);

                    email.Subject = $"ASSET MANAGEMENT - Return Approval Request - {gatepass_no}";

                    using (var client = new SmtpClient())
                    {
                        client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                        client.Authenticate(email_account, email_token);
                        client.Send(email);
                        client.Disconnect(true);
                    }

                    return "success";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string EMAIL_REQ_GATEPASS_RETURN_FINANCE_CONFIRMATION(int id_gatepass)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";
                string send_bcc = "leaantony1707@gmail.com";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, create_date, return_date, created_by, vendor_name, requestor_name FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_cc = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id='SESA100001'", conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    conn.Close();
                }

                if (send_to == "")
                {
                    return "No Finance Confirmation Email Found";
                }
                else
                {
                    string confirmationLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                    string emailContent = $@"
                        <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                            <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                                <div style='background-color: #007bff; padding: 20px; border-bottom: 2px solid #007bff;'>
                                    <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Return Finance Confirmation Request</h2>
                                </div>
                                <div style='padding: 25px;'>
                                    <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                    <p><strong>{requestor_name}</strong> has requested your finance confirmation for the return of Gate Pass <strong>{gatepass_no}</strong>. Please review the details below and confirm the return.</p>
                                    
                                    <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                        <table style='width: 100%; border-collapse: collapse;'>
                                            <tr style='background-color: #fcfcfc;'>
                                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                            </tr>
                                            <tr>
                                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                            </tr>
                                            <tr style='background-color: #fcfcfc;'>
                                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                            </tr>
                                            <tr>
                                                <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                                <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                                            </tr>
                                            <tr style='background-color: #fcfcfc;'>
                                                <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                                <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                            </tr>
                                        </table>
                                    </div>

                                    <div style='text-align: center; margin: 30px 0;'>
                                        <a href='{confirmationLink}' target='_blank' style='background-color: #007bff; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#007bff'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>Confirm Return</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                    </div>
                                    
                                    <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                    <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{confirmationLink}' style='color: #007bff; text-decoration: none;'>{confirmationLink}</a></p>
                                </div>
                                <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                    <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                    <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                                </div>
                            </div>
                        </div>";

                    var builder = new BodyBuilder();
                    builder.HtmlBody = emailContent;

                    email.Body = builder.ToMessageBody();

                    MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                    email.From.Add(from);

                    foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress to = new MailboxAddress("", email_to.Trim());
                        email.To.Add(to);
                    }

                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }

                    foreach (string email_bcc in send_bcc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress bcc = new MailboxAddress("", email_bcc.Trim());
                        email.Bcc.Add(bcc);
                    }

                    MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                    email.Bcc.Add(bcc1);

                    email.Subject = $"ASSET MANAGEMENT - Return Finance Confirmation Request - {gatepass_no}";

                    using (var client = new SmtpClient())
                    {
                        client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                        client.Authenticate(email_account, email_token);
                        client.Send(email);
                        client.Disconnect(true);
                    }

                    return "success";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string GATEPASS_APPROVED(int id_gatepass, string approver_name, string level, bool isDelegationInvolved = false, string originalApprover = "", string currentApproverSesaId = "")
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string type = "";
                string employee_name = "";
                string csr_to = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, type, employee_name, csr_to, create_date, return_date, created_by, vendor_name, requestor_name, location FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                                type = reader["type"].ToString() ?? "";
                                employee_name = reader["employee_name"].ToString() ?? "";
                                csr_to = reader["csr_to"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    if (isDelegationInvolved && !string.IsNullOrEmpty(originalApprover))
                    {
                        using (SqlCommand cmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = @originalApprover", conn))
                        {
                            cmd.Parameters.AddWithValue("@originalApprover", originalApprover);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string originalApproverEmail = reader["email"].ToString() ?? "";

                                    using (SqlCommand delegatedCmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = (SELECT delegated_to FROM mst_users WHERE sesa_id = @originalApprover)", conn))
                                    {
                                        delegatedCmd.Parameters.AddWithValue("@originalApprover", originalApprover);
                                        using (SqlDataReader delegatedReader = delegatedCmd.ExecuteReader())
                                        {
                                            if (delegatedReader.Read())
                                            {
                                                string delegatedEmail = delegatedReader["email"].ToString() ?? "";
                                                send_cc = originalApproverEmail + ";" + delegatedEmail;
                                            }
                                            else
                                            {
                                                send_cc = originalApproverEmail;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(currentApproverSesaId))
                        {
                            using (SqlCommand cmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = @currentApproverSesaId", conn))
                            {
                                cmd.Parameters.AddWithValue("@currentApproverSesaId", currentApproverSesaId);
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        send_cc = reader["email"].ToString() ?? "";
                                    }
                                }
                            }
                        }
                    }

                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string message = "";
                string tableRows = "";
                if (category == "REPAIR TO VENDOR" && level == "finance")
                {
                    message = $"Great news! Your assets for repair with gate pass order <strong>{gatepass_no}</strong> have been approved by <strong>{approver_name}</strong>. You can now proceed with the next steps.";
                    tableRows = $@"
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                        </tr>";
                }
                else if (category == "FA SALES - WITHOUT DISMANTLING")
                {
                    message = $"Great news! Your request for the Gate Pass of order <strong>{gatepass_no}</strong> has been approved by <strong>{approver_name}</strong>. You can now proceed with the next steps.";
                    string additionalField = "";
                    if (type == "VENDOR")
                    {
                        additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Vendor Name</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td></tr>";
                    }
                    else if (type == "EMPLOYEE")
                    {
                        additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Employee Name</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{employee_name}</td></tr>";
                    }
                    else
                    {
                        additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Destination To</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{csr_to}</td></tr>";
                    }
                    tableRows = $@"
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category} ({type})</td>
                        </tr>
                        {additionalField}";
                }
                else
                {
                    message = $"Great news! Your request for the Gate Pass of order <strong>{gatepass_no}</strong> has been approved by <strong>{approver_name}</strong>. You can now proceed with the next steps.";
                    tableRows = $@"
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                        </tr>";
                }

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Approval Approved</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>{message}</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        {tableRows}
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;

                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                if (!string.IsNullOrEmpty(send_cc))
                {
                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Approval Approved - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string GATEPASS_COMPLETED(int id_gatepass, string approver_name, string level, bool isDelegationInvolved = false, string originalApprover = "", string currentApproverSesaId = "")
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string type = "";
                string employee_name = "";
                string csr_to = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, type, employee_name, csr_to, create_date, return_date, created_by, vendor_name, requestor_name, location FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                                type = reader["type"].ToString() ?? "";
                                employee_name = reader["employee_name"].ToString() ?? "";
                                csr_to = reader["csr_to"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    if (isDelegationInvolved && !string.IsNullOrEmpty(originalApprover))
                    {
                        using (SqlCommand cmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = @originalApprover", conn))
                        {
                            cmd.Parameters.AddWithValue("@originalApprover", originalApprover);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string originalApproverEmail = reader["email"].ToString() ?? "";

                                    using (SqlCommand delegatedCmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = (SELECT delegated_to FROM mst_users WHERE sesa_id = @originalApprover)", conn))
                                    {
                                        delegatedCmd.Parameters.AddWithValue("@originalApprover", originalApprover);
                                        using (SqlDataReader delegatedReader = delegatedCmd.ExecuteReader())
                                        {
                                            if (delegatedReader.Read())
                                            {
                                                string delegatedEmail = delegatedReader["email"].ToString() ?? "";
                                                send_cc = originalApproverEmail + ";" + delegatedEmail;
                                            }
                                            else
                                            {
                                                send_cc = originalApproverEmail;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(currentApproverSesaId))
                        {
                            using (SqlCommand cmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = @currentApproverSesaId", conn))
                            {
                                cmd.Parameters.AddWithValue("@currentApproverSesaId", currentApproverSesaId);
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        send_cc = reader["email"].ToString() ?? "";
                                    }
                                }
                            }
                        }
                    }

                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string message = "";
                string tableRows = "";
                if (category == "REPAIR TO VENDOR" && level == "finance")
                {
                    message = $"Great news! Your assets for repair with gate pass order <strong>{gatepass_no}</strong> have been Completed by <strong>{approver_name}</strong>.";
                    tableRows = $@"
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                        </tr>";
                }
                else if (category == "FA SALES - WITHOUT DISMANTLING")
                {
                    message = $"Great news! Your request for the Gate Pass of order <strong>{gatepass_no}</strong> has been Completed by <strong>{approver_name}</strong>.";
                    string additionalField = "";
                    if (type == "VENDOR")
                    {
                        additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Vendor Name</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td></tr>";
                    }
                    else if (type == "EMPLOYEE")
                    {
                        additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Employee Name</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{employee_name}</td></tr>";
                    }
                    else
                    {
                        additionalField = $"<tr><td style='padding: 10px 15px; color: #666666;'>Destination To</td><td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{csr_to}</td></tr>";
                    }
                    tableRows = $@"
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category} ({type})</td>
                        </tr>
                        {additionalField}";
                }
                else
                {
                    message = $"Great news! Your request for the Gate Pass of order <strong>{gatepass_no}</strong> has been Completed by <strong>{approver_name}</strong>.";
                    tableRows = $@"
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                        </tr>
                        <tr>
                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                        </tr>
                        <tr style='background-color: #fcfcfc;'>
                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                        </tr>";
                }

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Completed</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>{message}</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        {tableRows}
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;

                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                if (!string.IsNullOrEmpty(send_cc))
                {
                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Gatepass Completed - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string GATEPASS_APPROVED_RETURN_SECURITY(int id_gatepass, string approver_name)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, create_date, return_date, created_by, vendor_name, requestor_name FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Return Approved</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>Great news! Your assets for repair with gate pass order <strong>{gatepass_no}</strong> have been returned and approved by <strong>{approver_name}</strong>. You can now proceed with the next steps.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;

                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Return Approved - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }


        public string GATEPASS_REJECTED(int id_gatepass)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, create_date, return_date, created_by, vendor_name, requestor_name FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #dc3545; padding: 20px; border-bottom: 2px solid #dc3545;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Approval Rejected</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>We regret to inform you that your request for the Gate Pass of order <strong>{gatepass_no}</strong> has been rejected. Please review the details below and consider revising your request if necessary.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #dc3545; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#dc3545'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;

                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Approval Rejected - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string GATEPASS_COMPLETE(int id_gatepass)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, create_date, return_date, created_by, vendor_name, requestor_name FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Approval Completed</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>Great news! Your approval request for the Gate Pass of order <strong>{gatepass_no}</strong> has been successfully completed. All necessary approvals have been obtained, and the process is now finalized.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;

                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Approval Completed - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
        public string GATEPASS_PIC_COMPLETE(int id_gatepass)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, create_date, return_date, created_by, vendor_name, requestor_name FROM v_gatepass WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id='SESA100001'", conn))
                    {
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                string sesa_email = reader["email"].ToString() ?? "";
                                send_to += ";" + sesa_email;
                            }
                        }
                    }
                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - PIC Completed</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>Great news! Your Gate Pass of order <strong>{gatepass_no}</strong> has been successfully completed with PIC confirmation.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;

                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - PIC Completed - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string EMAIL_REQ_NON_ASSET_GATEPASS(int id_gatepass, bool isDelegationInvolved = false, string originalApprover = "")
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";
                string send_bcc = "leaantony1707@gmail.com";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(@"
                SELECT h.gatepass_no, h.category, h.create_date, h.return_date, 
                       h.created_by, h.vendor_name, u.name AS requestor_name
                FROM tbl_gatepass_non_asset_header h
                LEFT JOIN mst_users u ON u.sesa_id = h.created_by
                WHERE h.id_gatepass = @id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 email FROM mst_users WHERE sesa_id=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        send_cc = cmd.ExecuteScalar()?.ToString() ?? "";
                    }

                    if (isDelegationInvolved && !string.IsNullOrEmpty(originalApprover))
                    {
                        using (SqlCommand cmd = new SqlCommand(@"
                            SELECT TOP 1 approval_by 
                            FROM tbl_gatepass_non_asset_approval 
                            WHERE id_gatepass=@id_gatepass AND approval_status='OPEN' 
                            ORDER BY approval_no", conn))
                        {
                            cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                            var nextApproverSesaId = cmd.ExecuteScalar()?.ToString() ?? "";

                            if (!string.IsNullOrEmpty(nextApproverSesaId))
                            {
                                using (SqlCommand checkDelegationCmd = new SqlCommand("SELECT delegated_to FROM mst_users WHERE sesa_id = @sesa_id", conn))
                                {
                                    checkDelegationCmd.Parameters.AddWithValue("@sesa_id", nextApproverSesaId);
                                    var delegatedTo = checkDelegationCmd.ExecuteScalar()?.ToString() ?? "";

                                    if (!string.IsNullOrEmpty(delegatedTo))
                                    {
                                        using (SqlCommand cmd2 = new SqlCommand("SELECT name, email FROM mst_users WHERE sesa_id = @original_approver OR sesa_id = @delegated_to", conn))
                                        {
                                            cmd2.Parameters.AddWithValue("@original_approver", nextApproverSesaId);
                                            cmd2.Parameters.AddWithValue("@delegated_to", delegatedTo);

                                            var names = new List<string>();
                                            var emails = new List<string>();

                                            using (SqlDataReader reader = cmd2.ExecuteReader())
                                            {
                                                while (reader.Read())
                                                {
                                                    names.Add(reader["name"].ToString() ?? "");
                                                    emails.Add(reader["email"].ToString() ?? "");
                                                }
                                            }

                                            if (emails.Count > 0)
                                            {
                                                send_to = string.Join(";", emails);
                                                send_to_name = string.Join(" & ", names);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        using (SqlCommand cmd2 = new SqlCommand("SELECT name, email FROM mst_users WHERE sesa_id = @sesa_id", conn))
                                        {
                                            cmd2.Parameters.AddWithValue("@sesa_id", nextApproverSesaId);
                                            using (SqlDataReader reader = cmd2.ExecuteReader())
                                            {
                                                if (reader.Read())
                                                {
                                                    send_to_name = reader["name"].ToString() ?? "";
                                                    send_to = reader["email"].ToString() ?? "";
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        using (SqlCommand cmd = new SqlCommand(@"
                            SELECT TOP 1 approval_by 
                            FROM tbl_gatepass_non_asset_approval 
                            WHERE id_gatepass=@id_gatepass AND approval_status='OPEN' 
                            ORDER BY approval_no", conn))
                        {
                            cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                            var nextApproverSesaId = cmd.ExecuteScalar()?.ToString() ?? "";

                            if (!string.IsNullOrEmpty(nextApproverSesaId))
                            {
                                using (SqlCommand checkDelegationCmd = new SqlCommand("SELECT delegated_to FROM mst_users WHERE sesa_id = @sesa_id", conn))
                                {
                                    checkDelegationCmd.Parameters.AddWithValue("@sesa_id", nextApproverSesaId);
                                    var delegatedTo = checkDelegationCmd.ExecuteScalar()?.ToString() ?? "";

                                    if (!string.IsNullOrEmpty(delegatedTo))
                                    {
                                        using (SqlCommand cmd2 = new SqlCommand("SELECT name, email FROM mst_users WHERE sesa_id = @original_approver OR sesa_id = @delegated_to", conn))
                                        {
                                            cmd2.Parameters.AddWithValue("@original_approver", nextApproverSesaId);
                                            cmd2.Parameters.AddWithValue("@delegated_to", delegatedTo);

                                            var names = new List<string>();
                                            var emails = new List<string>();

                                            using (SqlDataReader reader = cmd2.ExecuteReader())
                                            {
                                                while (reader.Read())
                                                {
                                                    names.Add(reader["name"].ToString() ?? "");
                                                    emails.Add(reader["email"].ToString() ?? "");
                                                }
                                            }

                                            if (emails.Count > 0)
                                            {
                                                send_to = string.Join(";", emails);
                                                send_to_name = string.Join(" & ", names);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        using (SqlCommand cmd2 = new SqlCommand("SELECT name, email FROM mst_users WHERE sesa_id = @sesa_id", conn))
                                        {
                                            cmd2.Parameters.AddWithValue("@sesa_id", nextApproverSesaId);
                                            using (SqlDataReader reader = cmd2.ExecuteReader())
                                            {
                                                if (reader.Read())
                                                {
                                                    send_to_name = reader["name"].ToString() ?? "";
                                                    send_to = reader["email"].ToString() ?? "";
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    conn.Close();
                }

                if (send_to == "")
                {
                    return "No Open Approval";
                }
                else
                {
                    string approvalLink = $"http://localhost:5242/Redirect/GatePassNonAssetList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                    string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #007bff; padding: 20px; border-bottom: 2px solid #007bff;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass Non-Asset - Approval Request</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p><strong>{requestor_name}</strong> has requested your approval for Gate Pass Non-Asset <strong>{gatepass_no}</strong>. Please review the details below.</p>
                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{approvalLink}' target='_blank' style='background-color: #007bff; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#007bff'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>Approve Gate Pass</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{approvalLink}' style='color: #007bff; text-decoration: none;'>{approvalLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                    var builder = new BodyBuilder();
                    builder.HtmlBody = emailContent;
                    email.Body = builder.ToMessageBody();

                    MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                    email.From.Add(from);

                    foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress to = new MailboxAddress("", email_to.Trim());
                        email.To.Add(to);
                    }

                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }

                    foreach (string email_bcc in send_bcc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress bcc = new MailboxAddress("", email_bcc.Trim());
                        email.Bcc.Add(bcc);
                    }

                    MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                    email.Bcc.Add(bcc1);

                    email.Subject = $"ASSET MANAGEMENT - Approval Request - {gatepass_no}";

                    using (var client = new SmtpClient())
                    {
                        client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                        client.Authenticate(email_account, email_token);
                        client.Send(email);
                        client.Disconnect(true);
                    }

                    return "success";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string EMAIL_REQ_NON_ASSET_GATEPASS_RETURN(int id_gatepass)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string create_date = "";
                string return_date = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";
                string send_bcc = "leaantony1707@gmail.com";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand("SELECT gatepass_no, category, create_date, return_date, created_by, vendor_name, requestor_name FROM v_gatepass_non_asset WHERE id_gatepass=@id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                create_date = reader["create_date"].ToString() ?? "";
                                return_date = reader["return_date"].ToString() ?? "-";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE [sesa_id]=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_cc = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id IN (SELECT TOP 1 approval_by FROM tbl_gatepass_non_asset_hod WHERE id_gatepass=@id_gatepass AND approval_status='OPEN')", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    conn.Close();
                }

                if (send_to == "")
                {
                    return "No Open Approval";
                }
                else
                {
                    string approvalLink = $"http://localhost:5242/Redirect/GatePassNonAssetList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                    string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #007bff; padding: 20px; border-bottom: 2px solid #007bff;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass Non-Asset - Return Approval Request</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p><strong>{requestor_name}</strong> has requested your approval for the return of Gate Pass Non-Asset <strong>{gatepass_no}</strong>. Please review the details below.</p>
                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Return Date</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{return_date}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{approvalLink}' target='_blank' style='background-color: #007bff; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#007bff'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>Approve Gate Pass Return</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{approvalLink}' style='color: #007bff; text-decoration: none;'>{approvalLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                    var builder = new BodyBuilder();
                    builder.HtmlBody = emailContent;

                    email.Body = builder.ToMessageBody();

                    MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                    email.From.Add(from);

                    foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress to = new MailboxAddress("", email_to.Trim());
                        email.To.Add(to);
                    }

                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }

                    foreach (string email_bcc in send_bcc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress bcc = new MailboxAddress("", email_bcc.Trim());
                        email.Bcc.Add(bcc);
                    }

                    MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                    email.Bcc.Add(bcc1);

                    email.Subject = $"ASSET MANAGEMENT - Return Approval Request - {gatepass_no}";

                    using (var client = new SmtpClient())
                    {
                        client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                        client.Authenticate(email_account, email_token);
                        client.Send(email);
                        client.Disconnect(true);
                    }
                    return "success";
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string GATEPASS_APPROVED_NON_ASSET(int id_gatepass, string approver_name, bool isDelegationInvolved = false, string originalApprover = "", string currentApproverSesaId = "")
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(@"
                SELECT h.gatepass_no, h.category, h.created_by, h.vendor_name, u.name AS requestor_name
                FROM tbl_gatepass_non_asset_header h
                LEFT JOIN mst_users u ON u.sesa_id = h.created_by
                WHERE h.id_gatepass = @id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    if (isDelegationInvolved && !string.IsNullOrEmpty(originalApprover))
                    {
                        using (SqlCommand cmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = @originalApprover", conn))
                        {
                            cmd.Parameters.AddWithValue("@originalApprover", originalApprover);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string originalApproverEmail = reader["email"].ToString() ?? "";
                                    using (SqlCommand delegatedCmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = (SELECT delegated_to FROM mst_users WHERE sesa_id = @originalApprover)", conn))
                                    {
                                        delegatedCmd.Parameters.AddWithValue("@originalApprover", originalApprover);
                                        using (SqlDataReader delegatedReader = delegatedCmd.ExecuteReader())
                                        {
                                            if (delegatedReader.Read())
                                            {
                                                string delegatedEmail = delegatedReader["email"].ToString() ?? "";
                                                send_cc = originalApproverEmail + ";" + delegatedEmail;
                                            }
                                            else
                                            {
                                                send_cc = originalApproverEmail;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(currentApproverSesaId))
                        {
                            using (SqlCommand cmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = @currentApproverSesaId", conn))
                            {
                                cmd.Parameters.AddWithValue("@currentApproverSesaId", currentApproverSesaId);
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        send_cc = reader["email"].ToString() ?? "";
                                    }
                                }
                            }
                        }
                    }

                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassNonAssetList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass Non-Asset - Approval Approved</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>Great news! Your request for the Gate Pass Non-Asset <strong>{gatepass_no}</strong> has been approved by <strong>{approver_name}</strong>. You can now proceed with the next steps.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;
                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                if (!string.IsNullOrEmpty(send_cc))
                {
                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Approval Approved - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string GATEPASS_COMPLETED_NON_ASSET(int id_gatepass, string approver_name, bool isDelegationInvolved = false, string originalApprover = "", string currentApproverSesaId = "")
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(@"
                SELECT h.gatepass_no, h.category, h.created_by, h.vendor_name, u.name AS requestor_name
                FROM tbl_gatepass_non_asset_header h
                LEFT JOIN mst_users u ON u.sesa_id = h.created_by
                WHERE h.id_gatepass = @id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    if (isDelegationInvolved && !string.IsNullOrEmpty(originalApprover))
                    {
                        using (SqlCommand cmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = @originalApprover", conn))
                        {
                            cmd.Parameters.AddWithValue("@originalApprover", originalApprover);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string originalApproverEmail = reader["email"].ToString() ?? "";
                                    using (SqlCommand delegatedCmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = (SELECT delegated_to FROM mst_users WHERE sesa_id = @originalApprover)", conn))
                                    {
                                        delegatedCmd.Parameters.AddWithValue("@originalApprover", originalApprover);
                                        using (SqlDataReader delegatedReader = delegatedCmd.ExecuteReader())
                                        {
                                            if (delegatedReader.Read())
                                            {
                                                string delegatedEmail = delegatedReader["email"].ToString() ?? "";
                                                send_cc = originalApproverEmail + ";" + delegatedEmail;
                                            }
                                            else
                                            {
                                                send_cc = originalApproverEmail;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(currentApproverSesaId))
                        {
                            using (SqlCommand cmd = new SqlCommand("SELECT email FROM mst_users WHERE sesa_id = @currentApproverSesaId", conn))
                            {
                                cmd.Parameters.AddWithValue("@currentApproverSesaId", currentApproverSesaId);
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    if (reader.Read())
                                    {
                                        send_cc = reader["email"].ToString() ?? "";
                                    }
                                }
                            }
                        }
                    }

                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassNonAssetList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass Non-Asset - Completed</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>Great news! Your request for the Gate Pass Non-Asset <strong>{gatepass_no}</strong> has been Completed by <strong>{approver_name}</strong>.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;
                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                if (!string.IsNullOrEmpty(send_cc))
                {
                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Gatepass Completed - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string GATEPASS_REJECTED_NON_ASSET(int id_gatepass)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(@"
                SELECT h.gatepass_no, h.category, h.created_by, h.vendor_name, u.name AS requestor_name
                FROM tbl_gatepass_non_asset_header h
                LEFT JOIN mst_users u ON u.sesa_id = h.created_by
                WHERE h.id_gatepass = @id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassNonAssetList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #dc3545; padding: 20px; border-bottom: 2px solid #dc3545;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass Non-Asset - Approval Rejected</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>We regret to inform you that your request for the Gate Pass Non-Asset <strong>{gatepass_no}</strong> has been rejected. Please review the details below and consider revising your request if necessary.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #dc3545; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#dc3545'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;
                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Approval Rejected - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string GATEPASS_COMPLETE_NON_ASSET(int id_gatepass)
        {
            try
            {
                var email = new MimeMessage();
                string gatepass_no = "";
                string category = "";
                string created_by = "";
                string vendor_name = "";
                string requestor_name = "";
                string send_to_name = "";
                string send_to = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();
                    using (SqlCommand cmd = new SqlCommand(@"
                        SELECT h.gatepass_no, h.category, h.created_by, h.vendor_name, u.name AS requestor_name
                        FROM tbl_gatepass_non_asset_header h
                        LEFT JOIN mst_users u ON u.sesa_id = h.created_by
                        WHERE h.id_gatepass = @id_gatepass", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_gatepass", id_gatepass);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                category = reader["category"].ToString() ?? "";
                                created_by = reader["created_by"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }
                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/GatePassNonAssetList?gatepass_no={Uri.EscapeDataString(gatepass_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass Non-Asset - Approval Completed</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>Great news! Your approval request for the Gate Pass Non-Asset <strong>{gatepass_no}</strong> has been successfully completed. All necessary approvals have been obtained, and the process is now finalized.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Category</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{category}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;
                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Approval Completed - {gatepass_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string EMAIL_REQ_INVOICE_APPROVAL(int id_invoice)
        {
            try
            {
                var email = new MimeMessage();
                string invoice_no = "";
                string gatepass_no = "";
                string vendor_name = "";
                string created_by = "";
                string requestor_name = "";
                string approval_sesa = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc_name = "";
                string send_cc = "";
                string send_bcc = "leaantony1707@gmail.com";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(@"
                        SELECT h.invoice_no, h.created_by, h.approval_sesa, 
                               g.gatepass_no, g.vendor_name, g.requestor_name
                        FROM tbl_gatepass_invoice_header h
                        LEFT JOIN v_gatepass g ON h.id_gatepass = g.id_gatepass
                        WHERE h.id_invoice = @id_invoice", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_invoice", id_invoice);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                invoice_no = reader["invoice_no"].ToString() ?? "";
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                created_by = reader["created_by"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                                approval_sesa = reader["approval_sesa"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_cc_name = reader["name"].ToString() ?? "";
                                send_cc = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@approval_sesa", conn))
                    {
                        cmd.Parameters.AddWithValue("@approval_sesa", approval_sesa);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    conn.Close();
                }

                if (string.IsNullOrEmpty(send_to))
                {
                    return "No approval email found";
                }

                string approvalLink = $"http://localhost:5242/Redirect/InvoiceApprovalList?invoice_no={Uri.EscapeDataString(invoice_no)}";

                string emailContent = $@"
                <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                    <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                        {GetLogoHtml()}
                        <div style='background-color: #007bff; padding: 20px; border-bottom: 2px solid #007bff;'>
                            <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Invoice Approval Request</h2>
                        </div>
                        <div style='padding: 25px;'>
                            <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                            <p><strong>{send_cc_name}</strong> has requested your approval for Invoice <strong>{invoice_no}</strong>. Please review the details below.</p>
                
                            <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                <table style='width: 100%; border-collapse: collapse;'>
                                    <tr style='background-color: #fcfcfc;'>
                                        <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Invoice No</td>
                                        <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{invoice_no}</td>
                                    </tr>
                                    <tr>
                                        <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Gate Pass No</td>
                                        <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                    </tr>
                                    <tr style='background-color: #fcfcfc;'>
                                        <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Vendor Name</td>
                                        <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{vendor_name}</td>
                                    </tr>
                                    <tr>
                                        <td style='padding: 10px 15px; color: #666666;'>Requestor Name</td>
                                        <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{requestor_name}</td>
                                    </tr>
                                </table>
                            </div>

                            <div style='text-align: center; margin: 30px 0;'>
                                <a href='{approvalLink}' target='_blank' style='background-color: #007bff; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#007bff'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>Approve Invoice</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                            </div>
                
                            <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                            <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{approvalLink}' style='color: #007bff; text-decoration: none;'>{approvalLink}</a></p>
                        </div>
                        <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                            <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                            <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                        </div>
                    </div>
                </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;
                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                if (!string.IsNullOrEmpty(send_cc))
                {
                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }
                }

                foreach (string email_bcc in send_bcc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress bcc = new MailboxAddress("", email_bcc.Trim());
                    email.Bcc.Add(bcc);
                }

                email.Subject = $"ASSET MANAGEMENT - Invoice Approval Request - {invoice_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string EMAIL_INVOICE_APPROVED(int id_invoice, string approver_name)
        {
            try
            {
                var email = new MimeMessage();
                string invoice_no = "";
                string gatepass_no = "";
                string vendor_code = "";
                string vendor_name = "";
                string vendor_email = "";
                string created_by = "";
                string requestor_name = "";
                string employee_id = "";
                string employee_name = "";
                string approval_sesa = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc_name = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(@"
                        SELECT h.invoice_no, h.created_by, h.approval_sesa,
                               g.gatepass_no, g.vendor_code, g.vendor_name, g.recipient_email AS vendor_email, g.requestor_name, g.employee_id, g.employee_name
                        FROM tbl_gatepass_invoice_header h
                        LEFT JOIN v_gatepass g ON h.id_gatepass = g.id_gatepass
                        WHERE h.id_invoice = @id_invoice", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_invoice", id_invoice);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                invoice_no = reader["invoice_no"].ToString() ?? "";
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                vendor_code = reader["vendor_code"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                vendor_email = reader["vendor_email"].ToString() ?? "";
                                created_by = reader["created_by"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                                employee_id = reader["employee_id"].ToString() ?? "";
                                employee_name = reader["employee_name"].ToString() ?? "";
                                approval_sesa = reader["approval_sesa"].ToString() ?? "";
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(employee_id))
                    {
                        using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@employee_id", conn))
                        {
                            cmd.Parameters.AddWithValue("@employee_id", employee_id);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    send_to_name = reader["name"].ToString() ?? "";
                                    send_to = reader["email"].ToString() ?? "";
                                }
                            }
                        }
                    }
                    else if (!string.IsNullOrEmpty(vendor_email))
                    {
                        send_to_name = vendor_name;
                        send_to = vendor_email;
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@approval_sesa", conn))
                    {
                        cmd.Parameters.AddWithValue("@approval_sesa", approval_sesa);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_cc_name = reader["name"].ToString() ?? "";
                                send_cc = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    string requestor_email = "";
                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 email FROM mst_users WHERE sesa_id=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        requestor_email = cmd.ExecuteScalar()?.ToString() ?? "";
                    }

                    if (!string.IsNullOrEmpty(send_cc) && !string.IsNullOrEmpty(requestor_email))
                    {
                        send_cc += ";" + requestor_email;
                    }
                    else if (!string.IsNullOrEmpty(requestor_email))
                    {
                        send_cc = requestor_email;
                    }

                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/InvoiceList?invoice_no={Uri.EscapeDataString(invoice_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #28a745; padding: 20px; border-bottom: 2px solid #28a745;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Invoice Approval Approved</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>Great news! Your invoice <strong>{invoice_no}</strong> has been approved by <strong>{send_cc_name}</strong>. You can now proceed with the next steps.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Invoice No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{invoice_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #28a745; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#28a745'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;

                try
                {
                    var exportController = new ExportController();
                    byte[] pdfBytes = exportController.GenerateInvoicePDFBytes(id_invoice);

                    var pdfMs = new MemoryStream(pdfBytes);
                    var pdfAttachment = new MimePart("application", "pdf")
                    {
                        Content = new MimeContent(pdfMs),
                        ContentDisposition = new ContentDisposition(ContentDisposition.Attachment),
                        ContentTransferEncoding = ContentEncoding.Base64,
                        FileName = $"Invoice_{invoice_no}.pdf"
                    };

                    builder.Attachments.Add(pdfAttachment);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error generating PDF: {ex.Message}");
                }

                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                if (!string.IsNullOrEmpty(send_cc))
                {
                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Invoice Approval Approved - {invoice_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string EMAIL_INVOICE_REJECTED(int id_invoice)
        {
            try
            {
                var email = new MimeMessage();
                string invoice_no = "";
                string gatepass_no = "";
                string vendor_name = "";
                string created_by = "";
                string requestor_name = "";
                string approval_sesa = "";
                string send_to_name = "";
                string send_to = "";
                string send_cc_name = "";
                string send_cc = "";

                using (SqlConnection conn = new SqlConnection(DbConnection()))
                {
                    conn.Open();

                    using (SqlCommand cmd = new SqlCommand(@"
                        SELECT h.invoice_no, h.created_by, h.approval_sesa,
                               g.gatepass_no, g.vendor_name, g.requestor_name
                        FROM tbl_gatepass_invoice_header h
                        LEFT JOIN v_gatepass g ON h.id_gatepass = g.id_gatepass
                        WHERE h.id_invoice = @id_invoice", conn))
                    {
                        cmd.Parameters.AddWithValue("@id_invoice", id_invoice);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                invoice_no = reader["invoice_no"].ToString() ?? "";
                                gatepass_no = reader["gatepass_no"].ToString() ?? "";
                                vendor_name = reader["vendor_name"].ToString() ?? "";
                                created_by = reader["created_by"].ToString() ?? "";
                                requestor_name = reader["requestor_name"].ToString() ?? "";
                                approval_sesa = reader["approval_sesa"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@created_by", conn))
                    {
                        cmd.Parameters.AddWithValue("@created_by", created_by);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_to_name = reader["name"].ToString() ?? "";
                                send_to = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    using (SqlCommand cmd = new SqlCommand("SELECT TOP 1 name, email FROM mst_users WHERE sesa_id=@approval_sesa", conn))
                    {
                        cmd.Parameters.AddWithValue("@approval_sesa", approval_sesa);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                send_cc_name = reader["name"].ToString() ?? "";
                                send_cc = reader["email"].ToString() ?? "";
                            }
                        }
                    }

                    conn.Close();
                }

                string viewLink = $"http://localhost:5242/Redirect/InvoiceList?invoice_no={Uri.EscapeDataString(invoice_no)}";

                string emailContent = $@"
                    <div style='font-family: Arial, sans-serif; font-size: 14px; color: #333333; line-height: 1.6; background-color: #ffffff;'>
                        <div style='max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff; overflow: hidden;'>
                            {GetLogoHtml()}
                            <div style='background-color: #dc3545; padding: 20px; border-bottom: 2px solid #dc3545;'>
                                <h2 style='margin: 0; color: #ffffff; font-size: 20px; text-align: center;'>Gate Pass - Invoice Approval Rejected</h2>
                            </div>
                            <div style='padding: 25px;'>
                                <p style='margin-top: 0;'>Dear <strong>{send_to_name}</strong>,</p>
                                <p>We regret to inform you that your invoice <strong>{invoice_no}</strong> has been rejected by <strong>{send_cc_name}</strong>. Please review the details below and consider revising your request if necessary.</p>
                                
                                <div style='margin: 20px 0; border: 1px solid #eeeeee; border-radius: 5px; overflow: hidden;'>
                                    <table style='width: 100%; border-collapse: collapse;'>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee; width: 40%;'>Invoice No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{invoice_no}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Gate Pass No</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{gatepass_no}</td>
                                        </tr>
                                        <tr style='background-color: #fcfcfc;'>
                                            <td style='padding: 10px 15px; color: #666666; border-bottom: 1px solid #eeeeee;'>Vendor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; border-bottom: 1px solid #eeeeee; color: #333333;'>{vendor_name}</td>
                                        </tr>
                                        <tr>
                                            <td style='padding: 10px 15px; color: #666666;'>Requestor Name</td>
                                            <td style='padding: 10px 15px; font-weight: bold; color: #333333;'>{requestor_name}</td>
                                        </tr>
                                    </table>
                                </div>

                                <div style='text-align: center; margin: 30px 0;'>
                                    <a href='{viewLink}' target='_blank' style='background-color: #dc3545; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 16px; display: inline-block; mso-padding-alt:0;text-underline-color:#dc3545'><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%;mso-text-raise:30pt'>&nbsp;</i><![endif]--><span style='mso-text-raise:15pt;'>View Details</span><!--[if mso]><i style='letter-spacing: 25px;mso-font-width:-100%'>&nbsp;</i><![endif]--></a>
                                </div>
                                
                                <p style='font-size: 12px; color: #999999; margin-bottom: 0;'>If the button above does not work, please copy and paste this link into your browser:</p>
                                <p style='font-size: 12px; margin-top: 5px; word-break: break-all;'><a href='{viewLink}' style='color: #007bff; text-decoration: none;'>{viewLink}</a></p>
                            </div>
                            <div style='background-color: #f8f8f8; padding: 15px; text-align: center; font-size: 12px; color: #888888; border-top: 1px solid #e0e0e0;'>
                                <p style='margin: 0; font-weight: bold;'>=== Please do not reply to this email ===</p>
                                <p style='margin: 5px 0 0;'>ASSET MANAGEMENT System</p>
                            </div>
                        </div>
                    </div>";

                var builder = new BodyBuilder();
                builder.HtmlBody = emailContent;
                email.Body = builder.ToMessageBody();

                MailboxAddress from = new MailboxAddress("ASSET MANAGEMENT", email_account);
                email.From.Add(from);

                foreach (string email_to in send_to.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    MailboxAddress to = new MailboxAddress("", email_to.Trim());
                    email.To.Add(to);
                }

                if (!string.IsNullOrEmpty(send_cc))
                {
                    foreach (string email_cc in send_cc.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        MailboxAddress cc = new MailboxAddress("", email_cc.Trim());
                        email.Cc.Add(cc);
                    }
                }

                MailboxAddress bcc1 = new MailboxAddress("", "leaantony1707@gmail.com");
                email.Bcc.Add(bcc1);

                email.Subject = $"ASSET MANAGEMENT - Invoice Approval Rejected - {invoice_no}";

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);
                    client.Authenticate(email_account, email_token);
                    client.Send(email);
                    client.Disconnect(true);
                }

                return "success";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

    }
}
