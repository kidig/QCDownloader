/*
 * Created by SharpDevelop.
 * User: kidig
 * Date: 06.11.2013
 * Time: 15:22
 * 
 */
using System;
using System.IO;
using System.Text;
using System.Configuration;
using TDAPIOLELib;
	

namespace QCDownloader
{
	class Program
	{

		public static void Main(string[] args)
		{
			Console.WriteLine("QC Requirements Downloader");
			
			ExportRequirements();
									
			Console.Write("Press any key to continue . . . ");
			Console.ReadKey(true);
		}

		
		private static void ExportRequirements()
		{
			string server_url = ConfigurationManager.AppSettings["SERVER_URL"];
			string username = ConfigurationManager.AppSettings["USER_NAME"];
			string password = ConfigurationManager.AppSettings["PASSWORD"];
			string domain = ConfigurationManager.AppSettings["DOMAIN"];
			string project = ConfigurationManager.AppSettings["PROJECT"];
			
			string req_file = ConfigurationManager.AppSettings["REQUIREMENTS_FILE"];
			string att_file = ConfigurationManager.AppSettings["ATTACHMENTS_FILE"];
			string att_path = ConfigurationManager.AppSettings["ATTACHMENTS_PATH"];
			
			if (!Directory.Exists(att_path))
			{
				Directory.CreateDirectory(att_path);
			}
			
			TDConnection tdc = new TDConnection();
			tdc.InitConnectionEx(server_url);
			tdc.ConnectProjectEx(domain, project, username, password);
			
			Console.WriteLine("Connected to QC Server");
			
			ReqFactory req_factory = (ReqFactory)tdc.ReqFactory;
			TDFilter req_filter = (TDFilter)req_factory.Filter;
			
			
			/**
 			 *  Set your own filters for requirements below
			 */
			
			// req_filter["RQ_REQ_PATH"] = "AAAAAGAAE*";
			
			
			StreamWriter rfs = new StreamWriter(File.Open(req_file, FileMode.Create),
			                                           Encoding.Default,
			                                           1024);
			
			StreamWriter afs = new StreamWriter(File.Open(att_file, FileMode.Create),
			                                           Encoding.Default,
			                                           1024);
			
			foreach (Req r in req_filter.NewList()) 
			{
				string name = r.Name.Replace("\"","").Replace("\t","").Trim();
								
				Console.WriteLine("Req \"{0}\"", name);
				
				rfs.WriteLine(String.Join("\t", new String[]{
				                          	r.ID.ToString(),
				                          	r.ParentId.ToString(),
				                          	r["RQ_REQ_PATH"],
				                          	name,
				                          	r["RQ_VTS"].ToString()
				                          }));
				
				if (!r.HasAttachment)
					continue;
				
				AttachmentFactory att_factory = r.Attachments;
				
				foreach (Attachment a in att_factory.NewList(""))
				{
					Console.WriteLine("Attachment \"{0}\"", a.Name);
					
					afs.WriteLine(String.Join("\t", new String[]{
					                          	r.ID.ToString(),
					                          	a.ID.ToString(),
					                          	a.Name,
					                          	a.FileSize.ToString(),
					                          	a.LastModified.ToShortDateString()
					                          }));
					
					IExtendedStorage storage = a.AttachmentStorage;
					storage.ClientPath = Path.GetFullPath(att_path) + "\\";
					storage.Load(a.Name, true);
				}
				
			} 
			
			rfs.Close();
			afs.Close();
			
			tdc.Disconnect();
			tdc.Logout();
			
			Console.WriteLine("Disconnected.");			
		}
	}
}