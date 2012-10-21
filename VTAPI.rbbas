#tag Module
Protected Module VTAPI
	#tag Method, Flags = &h0
		Function AddComment(ResourceID As String, APIKey As String, Comment As String) As JSONItem
		  Dim js As New JSONItem
		  js.Value("apikey") = APIKey
		  js.Value("resource") = ResourceID
		  js.Value("comment") = Comment
		  Return SendRequest(CommentPut_URL, js)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CRLF() As String
		  Return Encodings.ASCII.Chr(13) + Encodings.ASCII.Chr(10)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetReport(ResourceID As String, APIKey As String, ReportType As Integer = 0) As JSONItem
		  'Report type can be 0 or 1. 0 is for files, 1 is for URLs.
		  Dim js As New JSONItem
		  js.Value("apikey") = APIKey
		  
		  Select Case ReportType
		  Case TypeFile
		    js.Value("resource") = ResourceID
		    Return SendRequest(FileReport_URL, js)
		  Case TypeURL
		    js.Value("url") = ResourceID
		    Return SendRequest(URLReport_URL, js)
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Rescan(ResourceID As String, APIKey As String) As JSONItem
		  Dim js As New JSONItem
		  js.Value("resource") =ResourceID
		  js.Value("apikey") = APIKey
		  Return SendRequest(FileRescan_URL, js)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SendRequest(URL As String, Request As JSONItem, VTSock As HTTPSecureSocket = Nil) As JSONItem
		  If VTSock = Nil Then VTSock = New HTTPSecureSocket
		  VTSock.Secure = True
		  dim formData As New Dictionary
		  For Each Name As String In Request.Names
		    formData.Value(Name) = Request.Value(Name)
		  Next
		  VTSock.SetFormData(formData)
		  Dim s As String = VTSock.Post(URL, 5)
		  Return New JSONItem(s)
		End Function
	#tag EndMethod


	#tag Constant, Name = CommentPut_URL, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/comments/put", Scope = Private
	#tag EndConstant

	#tag Constant, Name = FileReport_URL, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/file/report", Scope = Private
	#tag EndConstant

	#tag Constant, Name = FileRescan_URL, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/file/rescan", Scope = Private
	#tag EndConstant

	#tag Constant, Name = FileSubmit_URL, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/file/scan", Scope = Private
	#tag EndConstant

	#tag Constant, Name = TypeFile, Type = Double, Dynamic = False, Default = \"0", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = TypeURL, Type = Double, Dynamic = False, Default = \"1", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = URLReport_URL, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/url/scan", Scope = Private
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			InheritedFrom="Object"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			InheritedFrom="Object"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
