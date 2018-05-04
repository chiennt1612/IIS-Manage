Imports Microsoft.Web.Administration
namespace IISManage
	public module Utils
		'-- c. Folder and File
		Private Function GetSizeFolder(Byval sourcePath As String) As Double
			Dim folderSize As Double
			folderSize = 0
			Try
				If(Directory.Exists(sourcePath)) Then
					Dim source As DirectoryInfo
					Dim dir As DirectoryInfo, file As FileInfo
					source = New DirectoryInfo (sourcePath)
					For Each file In source.GetFiles()
						folderSize += file.Length
					Next
					For Each dir In source.GetDirectories()
						folderSize += GetSizeFolder(sourcePath & "\" & dir.Name)
					Next
				End If
			Catch
			End Try
			Return folderSize
		End Function
		Private Sub CopyFolder(Byval sourcePath As String, Byval targetPath As String)
			Dim source As DirectoryInfo, target As DirectoryInfo
			Dim dir As DirectoryInfo, file As FileInfo
			source = New DirectoryInfo (sourcePath)
			If (Not Directory.Exists(targetPath)) Then Directory.CreateDirectory (targetPath)
			target = New DirectoryInfo (targetPath)
			For Each dir In source.GetDirectories()
				CopyFolder(sourcePath & "\" & dir.Name, targetPath & "\" & dir.Name)
			Next
			For Each file In source.GetFiles()
				file.CopyTo(Path.Combine(target.FullName, file.Name))
			Next
		End Sub
		Private Sub SetWebConfig(Byval sourcePath As String, Byval targetPath As String, Byval EncryptConnStr As String)
			Dim txtConfig As String
			txtConfig = File.ReadAllText (sourcePath)
			txtConfig = Replace(txtConfig, "[=DBCONNSTR]", EncryptConnStr)
			File.WriteAllText (targetPath, txtConfig)
		End Sub
		Public Function GetSizeFolderByCompany(Byval CompanyCode As String) As Double
			Return GetSizeFolder(REPLACE(CompanyCodePath, "[COMPANYSCHEMA]", CompanyCode))
		End Function
		Public Sub Test(Byval sourcePath As String, Byval targetPath As String, Byval EncryptConnStr As String)
			CopyFolder(sourcePath, targetPath)
			SetWebConfig(sourcePath & "\web.config", targetPath & "\web.config", EncryptConnStr)
		End Sub
		Public Function FileReadAllText(Byval sourcePath As String) As String
			FileReadAllText = File.ReadAllText (sourcePath)
		End Function
		'-- e. Create website
		Private Function CreateNewApplicationPoolWithUsername(Byval poolName As String, Byval poolUserID as String, Byval poolPass as String) As String
			Dim CreateReturn As String		
			CreateReturn = "-1"
			poolName = Trim(poolName)
			Try
				Dim server As New ServerManager, pool As ApplicationPool, pools As ApplicationPoolCollection, kt As Boolean
				pools = server.ApplicationPools
				pool = nothing
				kt = False
				For Each pool in pools
					If (Trim(pool.Name) = poolName) Then Kt = True
				Next
				If (Kt) Then
					' Pool đã có
					CreateReturn = "-8"
				Else
					pool = pools.Add(poolName)
					If (Not pool Is Nothing) Then
						pool.AutoStart = true
						CreateReturn = pool.AutoStart & ","
						CreateReturn = CreateReturn & pool.ManagedRuntimeVersion & ","
						CreateReturn = CreateReturn & pool.Name
						'CreateReturn = "1"
						If(poolUserID <> "" AND poolPass <> "") Then
							pool.AutoStart = True
							pool.ProcessModel.IdentityType = ProcessModelIdentityType.SpecificUser
							pool.ProcessModel.UserName = poolUserID
							pool.ProcessModel.Password = poolPass
						End If
						server.CommitChanges()
					Else 
						' Không tạo đc
						CreateReturn = "-7"
					End If
				End If
				pool = nothing
				pools = nothing
				server = nothing		
			Catch
				' Lỗi try catch
				CreateReturn = "-1"
			End Try
			Return CreateReturn
		End Function
		Private Function CreateNewApplicationPool(Byval poolName As String) As String
			Return CreateNewApplicationPoolWithUsername(poolName, "", "")
		End Function
		Private Function DeleteWebsite(Byval WebName As String) As String
			Try
				Dim server1 As New ServerManager, sites1 As SiteCollection, s1 As Site
				sites1 = server1.Sites
				s1 = sites1(WebName)
				sites1.Remove(s1)
				server1.CommitChanges()
				s1 = nothing
				sites1 = nothing
				server1 = nothing
				Return "1.Thành công"
			Catch E As Exception
				Return "-11." & E.ToString
			End Try
		End Function
		Public Function StopWebsite(Byval WebName As String) As String
			Try
				Dim server1 As New ServerManager, sites1 As SiteCollection, s1 As Site
				sites1 = server1.Sites
				s1 = sites1(WebName)
				s1.Stop()
				server1.CommitChanges()
				s1 = nothing
				sites1 = nothing
				server1 = nothing
				Return "1.Thành công"
			Catch E As Exception
				Return "-11." & E.ToString
			End Try
		End Function
		Public Function StartWebsite(Byval WebName As String) As String
			Try
				Dim server1 As New ServerManager, sites1 As SiteCollection, s1 As Site
				sites1 = server1.Sites
				s1 = sites1(WebName)
				s1.Start()
				server1.CommitChanges()
				s1 = nothing
				sites1 = nothing
				server1 = nothing
				Return "1.Thành công"
			Catch E As Exception
				Return "-11." & E.ToString
			End Try
		End Function
		Private Function CreateNewWebsite(Byval WebName As String, Byval path as String, Byval port As String, Byval hostName As String) As String
			Return CreateNewWebsiteWithPool(WebName, path, port, hostName, "")
		End Function
		Public Function CreateNewWebsiteWithPool(Byval WebName As String, Byval path as String, Byval port As String, Byval hostName As String, Byval poolName as String) As String
			Dim CreateReturn As String		
			CreateReturn = "-1.Lỗi"
			WebName = Trim(WebName)
			path = Trim(path)
			port = Trim(port)
			hostName = Trim(hostName)
			poolName = Trim(poolName)
			Try
				'Dim usr As String, pwd As String
				'ReadLoginIIS(usr, pwd)
				'CreateNewApplicationPoolWithUsername(poolName, usr, pwd)
				CreateNewApplicationPool(poolName)
			Catch
			End Try
			Try
				DeleteWebsite (WebName)
			Catch
			End Try
			Try
				Dim server As New ServerManager, sites As SiteCollection, s As Site, bindingInfo As String, ip As String
				sites = server.Sites
				s = nothing
				ip = "*"
				bindingInfo = string.Format("{0}:{1}:{2}", ip, port, hostName)
				s = sites.Add(WebName, "http", bindingInfo, path)

				If (Not s Is Nothing) Then
					'CreateReturn = poolName & ","
					'CreateReturn = CreateReturn & WebName & ","
					'CreateReturn = CreateReturn & hostName & ","
					'CreateReturn = CreateReturn & ip & ","
					'CreateReturn = CreateReturn & port & ","
					'CreateReturn = CreateReturn & path
					CreateReturn = "1.Thành công"
					If (poolName <> "") Then s.ApplicationDefaults.ApplicationPoolName = poolName
					server.CommitChanges()
				Else 
					' Không tạo đc
					CreateReturn = "-7.Không tạo được Website"
				End If
				
				s = nothing
				sites = nothing
				server = nothing		
			Catch E As Exception
				' Lỗi try catch
				CreateReturn = "-1." & E.ToString
			End Try
			Return CreateReturn
		End Function
		Public Function GetWebList(Byval Token as String) As String
			Dim StrListWebsite As String		
			StrListWebsite = ""
			Try
				If(CheckToken(Token, 0, 0))Then
					Dim server As New ServerManager 
					Dim sites As SiteCollection, s As Site, d As ApplicationDefaults
					sites = server.Sites
					For Each s in sites	
						'StrListWebsite = StrListWebsite & ";"
						'StrListWebsite = StrListWebsite & s.Id & ","
						'StrListWebsite = StrListWebsite & s.State & ","
						'StrListWebsite = StrListWebsite & s.Name & ","
						d = s.ApplicationDefaults
						'StrListWebsite = StrListWebsite & d.ApplicationPoolName & ",("
						'StrListWebsite = StrListWebsite & d.Schema.ToString & ""
						Dim dirs As VirtualDirectoryCollection, dir As VirtualDirectory, physicalPath As String, _
							apps As ApplicationCollection, app As Application
						apps = s.Applications
						app = apps(0)
						'For Each app in apps
							dirs = app.VirtualDirectories
							'StrListWebsite = StrListWebsite & "["
							dir = dirs(0)
							physicalPath = dir.PhysicalPath
							'For Each dir in dirs
							'	path = dir.Path
							'	physicalPath = dir.PhysicalPath
								'StrListWebsite = StrListWebsite & path & ","
								'StrListWebsite = StrListWebsite & physicalPath
							'Next
							'StrListWebsite = StrListWebsite & "]"						
						'Next
						'StrListWebsite = StrListWebsite & ")"
						StrListWebsite &= ",{""ID"":""" & s.Id & """,""Name"":""" & s.Name & """,""State"":""" & s.State & """,""ApplicationPoolName"":""" & d.ApplicationPoolName & """,""physicalPath"":""" & Replace(physicalPath, "\", "\\") & """}"
						dir = nothing
						dirs = nothing
						app = nothing
						apps = nothing
					Next
					If(Len(StrListWebsite)>1)Then StrListWebsite = "{""ResponseStatus"":""1"",""Message"":""Thành công"",""Data"":[" & Right(StrListWebsite, Len(StrListWebsite)-1) & "]}"
					d = nothing
					s = nothing
					sites = nothing
					server = nothing
				Else
					StrListWebsite = "{""ResponseStatus"":""-9"",""Message"":""Token không hợp lệ"",""Data"":[]}"
				End If
				
			Catch E As Exception
				StrListWebsite = "{""ResponseStatus"":""-99"",""Message"":""" & Replace(E.ToString, """", "\""") & """,""Data"":[]}"
			End Try
			Return StrListWebsite
		End Function
		Public Function GetApplicationPool(Byval Token As String) As String
			Dim StrListAppPool As String		
			StrListAppPool = ""
			Try
				If(CheckToken(Token, 0, 0))Then
					Dim server As New ServerManager 
					Dim pools As ApplicationPoolCollection, pool As ApplicationPool
					pools = server.ApplicationPools
					For Each pool in pools	
						'StrListAppPool = StrListAppPool & ";"
						'StrListAppPool = StrListAppPool & pool.AutoStart & ","
						'StrListAppPool = StrListAppPool & pool.ManagedRuntimeVersion & ","
						'StrListAppPool = StrListAppPool & pool.Name
						StrListAppPool &= ",{""Name"":""" & pool.Name & """,""ManagedRuntimeVersion"":""" & pool.ManagedRuntimeVersion & """,""AutoStart"":""" & pool.AutoStart & """}"
					Next
					If(Len(StrListAppPool)>1)Then StrListAppPool = "{""ResponseStatus"":""1"",""Message"":""Thành công"",""Data"":[" & Right(StrListAppPool, Len(StrListAppPool)-1) & "]}"
					pool = nothing
					pools = nothing
					server = nothing
				Else
					StrListAppPool = "{""ResponseStatus"":""-9"",""Message"":""Token không hợp lệ"",""Data"":[]}"
				End If
				
			Catch E As Exception
				StrListAppPool = "{""ResponseStatus"":""-99"",""Message"":""" & Replace(E.ToString, """", "\""") & """,""Data"":[]}"
			End Try
			Return StrListAppPool
		End Function
	end module
End namespace
