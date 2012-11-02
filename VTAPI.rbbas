#tag Module
Protected Module VTAPI
	#tag Method, Flags = &h0
		Function AddComment(ResourceID As String, APIKey As String, Comment As String) As JSONItem
		  Dim js As New JSONItem
		  js.Value("apikey") = APIKey
		  js.Value("resource") = ResourceID
		  js.Value("comment") = Comment
		  Return SendRequest(VT_Put_Comment, js)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ConstructUpload(File As FolderItem) As String
		  dim rawdata, data as string
		  dim ReadStream as BinaryStream
		  
		  ReadStream = BinaryStream.Open(File)
		  rawdata = ReadStream.Read(File.Length)
		  
		  
		  // Step 1: Initialize data string
		  
		  // start with boundary line
		  data="--" + MIMEBoundary + CRLF
		  
		  // Step 2: Add each of the text-based form fields
		  // as a separate MIME part, followed by a boundary
		  
		  // Now the file name of the first image
		  'data = data + "Content-Disposition: form-data; name=""file""" + CRLF + CRLF + File.Name + CRLF
		  
		  // boundary to separate the above from the next part
		  'data = data + "--" + MIMEBoundary + CRLF
		  
		  // Next comes the actual file itself, with headers indicating what type it is
		  // and what content type we&apos;re using (raw binary in this case)
		  
		  data = data + "Content-Disposition: form-data; name=""file""; filename=""" +  File.Name + """" + CRLF + "Content-Type: " + HTTPMimeString(File.Name) + CRLF + "Content-Length: " + str(lenb(rawdata)) + CRLF + "Content-Transfer-Encoding: binary" + CRLF + CRLF + rawdata + CRLF
		  
		  
		  // Now our closing boundary marker, and our data is ready to send
		  data = data + "--" + MIMEBoundary + CRLF + CRLF
		  //whew...
		  
		  Return data
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function CRLF() As String
		  Return Encodings.ASCII.Chr(13) + Encodings.ASCII.Chr(10)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function GetReport(ResourceID As String, APIKey As String, ReportType As ReportType) As JSONItem
		  'Report type can be 0 or 1. 0 is for files, 1 is for URLs.
		  Dim js As New JSONItem
		  js.Value("apikey") = APIKey
		  
		  Select Case ReportType
		  Case VTAPI.ReportType.FileReport
		    js.Value("resource") = ResourceID
		    Return SendRequest(VT_Get_File, js)
		  Case VTAPI.ReportType.URLReport
		    js.Value("url") = ResourceID
		    Return SendRequest(VT_Get_URL, js)
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function HTTPMimeString(FileName As String) As String
		  'This method is from here: https://github.com/bskrtich/RBHTTPServer
		  Dim ext As String = NthField(FileName, ".", CountFields(FileName, "."))
		  ext = Lowercase(ext)
		  
		  Select Case ext
		  Case "ez"
		    Return "application/andrew-inset"
		    
		  Case "aw"
		    Return "application/applixware"
		    
		  Case "atom"
		    Return "application/atom+xml"
		    
		  Case "atomcat"
		    Return "application/atomcat+xml"
		    
		  Case "atomsvc"
		    Return "application/atomsvc+xml"
		    
		  Case "ccxml"
		    Return "application/ccxml+xml"
		    
		  Case "cdmia"
		    Return "application/cdmi-capability"
		    
		  Case "cdmic"
		    Return "application/cdmi-container"
		    
		  Case "cdmid"
		    Return "application/cdmi-domain"
		    
		  Case "cdmio"
		    Return "application/cdmi-object"
		    
		  Case "cdmiq"
		    Return "application/cdmi-queue"
		    
		  Case "cu"
		    Return "application/cu-seeme"
		    
		  Case "davmount"
		    Return "application/davmount+xml"
		    
		  Case "dssc"
		    Return "application/dssc+der"
		    
		  Case "xdssc"
		    Return "application/dssc+xml"
		    
		  Case "ecma"
		    Return "application/ecmascript"
		    
		  Case "emma"
		    Return "application/emma+xml"
		    
		  Case "epub"
		    Return "application/epub+zip"
		    
		  Case "exi"
		    Return "application/exi"
		    
		  Case "pfr"
		    Return "application/font-tdpfr"
		    
		  Case "stk"
		    Return "application/hyperstudio"
		    
		  Case "ipfix"
		    Return "application/ipfix"
		    
		  Case "jar"
		    Return "application/java-archive"
		    
		  Case "ser"
		    Return "application/java-serialized-object"
		    
		  Case "class"
		    Return "application/java-vm"
		    
		  Case "js"
		    Return "application/javascript"
		    
		  Case "json"
		    Return "application/json"
		    
		  Case "lostxml"
		    Return "application/lost+xml"
		    
		  Case "hqx"
		    Return "application/mac-binhex40"
		    
		  Case "cpt"
		    Return "application/mac-compactpro"
		    
		  Case "mads"
		    Return "application/mads+xml"
		    
		  Case "mrc"
		    Return "application/marc"
		    
		  Case "mrcx"
		    Return "application/marcxml+xml"
		    
		  Case "ma", "nb", "mb"
		    Return "application/mathematica"
		    
		  Case "mathml"
		    Return "application/mathml+xml"
		    
		  Case "mbox"
		    Return "application/mbox"
		    
		  Case "mscml"
		    Return "application/mediaservercontrol+xml"
		    
		  Case "meta4"
		    Return "application/metalink4+xml"
		    
		  Case "mets"
		    Return "application/mets+xml"
		    
		  Case "mods"
		    Return "application/mods+xml"
		    
		  Case "m21", "mp21"
		    Return "application/mp21"
		    
		  Case "mp4s"
		    Return "application/mp4"
		    
		  Case "doc", "dot"
		    Return "application/msword"
		    
		  Case "mxf"
		    Return "application/mxf"
		    
		  Case "bin", "dms", "lha", "lrf", "lzh", "so", "iso", "dmg", "dist", "distz", "pkg", "bpk", "dump", "elc", "deploy", "mobipocket-ebook"
		    Return "application/octet-stream"
		    
		  Case "oda"
		    Return "application/oda"
		    
		  Case "opf"
		    Return "application/oebps-package+xml"
		    
		  Case "ogx"
		    Return "application/ogg"
		    
		  Case "onetoc", "onetoc2", "onetmp", "onepkg"
		    Return "application/onenote"
		    
		  Case "xer"
		    Return "application/patch-ops-error+xml"
		    
		  Case "pdf"
		    Return "application/pdf"
		    
		  Case "pgp"
		    Return "application/pgp-encrypted"
		    
		  Case "asc", "sig"
		    Return "application/pgp-signature"
		    
		  Case "prf"
		    Return "application/pics-rules"
		    
		  Case "p10"
		    Return "application/pkcs10"
		    
		  Case "p7m", "p7c"
		    Return "application/pkcs7-mime"
		    
		  Case "p7s"
		    Return "application/pkcs7-signature"
		    
		  Case "p8"
		    Return "application/pkcs8"
		    
		  Case "ac"
		    Return "application/pkix-attr-cert"
		    
		  Case "cer"
		    Return "application/pkix-cert"
		    
		  Case "crl"
		    Return "application/pkix-crl"
		    
		  Case "pkipath"
		    Return "application/pkix-pkipath"
		    
		  Case "pki"
		    Return "application/pkixcmp"
		    
		  Case "pls"
		    Return "application/pls+xml"
		    
		  Case "ai", "eps", "ps"
		    Return "application/postscript"
		    
		  Case "cww"
		    Return "application/prs.cww"
		    
		  Case "pskcxml"
		    Return "application/pskc+xml"
		    
		  Case "rdf"
		    Return "application/rdf+xml"
		    
		  Case "rif"
		    Return "application/reginfo+xml"
		    
		  Case "rnc"
		    Return "application/relax-ng-compact-syntax"
		    
		  Case "rl"
		    Return "application/resource-lists+xml"
		    
		  Case "rld"
		    Return "application/resource-lists-diff+xml"
		    
		  Case "rs"
		    Return "application/rls-services+xml"
		    
		  Case "rsd"
		    Return "application/rsd+xml"
		    
		  Case "rss"
		    Return "application/rss+xml"
		    
		  Case "rtf"
		    Return "application/rtf"
		    
		  Case "sbml"
		    Return "application/sbml+xml"
		    
		  Case "scq"
		    Return "application/scvp-cv-request"
		    
		  Case "scs"
		    Return "application/scvp-cv-response"
		    
		  Case "spq"
		    Return "application/scvp-vp-request"
		    
		  Case "spp"
		    Return "application/scvp-vp-response"
		    
		  Case "sdp"
		    Return "application/sdp"
		    
		  Case "setpay"
		    Return "application/set-payment-initiation"
		    
		  Case "setreg"
		    Return "application/set-registration-initiation"
		    
		  Case "shf"
		    Return "application/shf+xml"
		    
		  Case "smi", "smil"
		    Return "application/smil+xml"
		    
		  Case "rq"
		    Return "application/sparql-query"
		    
		  Case "srx"
		    Return "application/sparql-results+xml"
		    
		  Case "gram"
		    Return "application/srgs"
		    
		  Case "grxml"
		    Return "application/srgs+xml"
		    
		  Case "sru"
		    Return "application/sru+xml"
		    
		  Case "ssml"
		    Return "application/ssml+xml"
		    
		  Case "tei", "teicorpus"
		    Return "application/tei+xml"
		    
		  Case "tfi"
		    Return "application/thraud+xml"
		    
		  Case "tsd"
		    Return "application/timestamped-data"
		    
		  Case "plb"
		    Return "application/vnd.3gpp.pic-bw-large"
		    
		  Case "psb"
		    Return "application/vnd.3gpp.pic-bw-small"
		    
		  Case "pvb"
		    Return "application/vnd.3gpp.pic-bw-var"
		    
		  Case "tcap"
		    Return "application/vnd.3gpp2.tcap"
		    
		  Case "pwn"
		    Return "application/vnd.3m.post-it-notes"
		    
		  Case "aso"
		    Return "application/vnd.accpac.simply.aso"
		    
		  Case "imp"
		    Return "application/vnd.accpac.simply.imp"
		    
		  Case "acu"
		    Return "application/vnd.acucobol"
		    
		  Case "atc", "acutc"
		    Return "application/vnd.acucorp"
		    
		  Case "air"
		    Return "application/vnd.adobe.air-application-installer-package+zip"
		    
		  Case "fxp", "fxpl"
		    Return "application/vnd.adobe.fxp"
		    
		  Case "xdp"
		    Return "application/vnd.adobe.xdp+xml"
		    
		  Case "xfdf"
		    Return "application/vnd.adobe.xfdf"
		    
		  Case "ahead"
		    Return "application/vnd.ahead.space"
		    
		  Case "azf"
		    Return "application/vnd.airzip.filesecure.azf"
		    
		  Case "azs"
		    Return "application/vnd.airzip.filesecure.azs"
		    
		  Case "azw"
		    Return "application/vnd.amazon.ebook"
		    
		  Case "acc"
		    Return "application/vnd.americandynamics.acc"
		    
		  Case "ami"
		    Return "application/vnd.amiga.ami"
		    
		  Case "apk"
		    Return "application/vnd.android.package-archive"
		    
		  Case "cii"
		    Return "application/vnd.anser-web-certificate-issue-initiation"
		    
		  Case "fti"
		    Return "application/vnd.anser-web-funds-transfer-initiation"
		    
		  Case "atx"
		    Return "application/vnd.antix.game-component"
		    
		  Case "mpkg"
		    Return "application/vnd.apple.installer+xml"
		    
		  Case "m3u8"
		    Return "application/vnd.apple.mpegurl"
		    
		  Case "swi"
		    Return "application/vnd.aristanetworks.swi"
		    
		  Case "aep"
		    Return "application/vnd.audiograph"
		    
		  Case "mpm"
		    Return "application/vnd.blueice.multipass"
		    
		  Case "bmi"
		    Return "application/vnd.bmi"
		    
		  Case "rep"
		    Return "application/vnd.businessobjects"
		    
		  Case "cdxml"
		    Return "application/vnd.chemdraw+xml"
		    
		  Case "mmd"
		    Return "application/vnd.chipnuts.karaoke-mmd"
		    
		  Case "cdy"
		    Return "application/vnd.cinderella"
		    
		  Case "cla"
		    Return "application/vnd.claymore"
		    
		  Case "rp9"
		    Return "application/vnd.cloanto.rp9"
		    
		  Case "c4g", "c4d", "c4f", "c4p", "c4u"
		    Return "application/vnd.clonk.c4group"
		    
		  Case "c11amc"
		    Return "application/vnd.cluetrust.cartomobile-config"
		    
		  Case "c11amz"
		    Return "application/vnd.cluetrust.cartomobile-config-pkg"
		    
		  Case "csp"
		    Return "application/vnd.commonspace"
		    
		  Case "cdbcmsg"
		    Return "application/vnd.contact.cmsg"
		    
		  Case "cmc"
		    Return "application/vnd.cosmocaller"
		    
		  Case "clkx"
		    Return "application/vnd.crick.clicker"
		    
		  Case "clkk"
		    Return "application/vnd.crick.clicker.keyboard"
		    
		  Case "clkp"
		    Return "application/vnd.crick.clicker.palette"
		    
		  Case "clkt"
		    Return "application/vnd.crick.clicker.template"
		    
		  Case "clkw"
		    Return "application/vnd.crick.clicker.wordbank"
		    
		  Case "wbs"
		    Return "application/vnd.criticaltools.wbs+xml"
		    
		  Case "pml"
		    Return "application/vnd.ctc-posml"
		    
		  Case "ppd"
		    Return "application/vnd.cups-ppd"
		    
		  Case "car"
		    Return "application/vnd.curl.car"
		    
		  Case "pcurl"
		    Return "application/vnd.curl.pcurl"
		    
		  Case "rdz"
		    Return "application/vnd.data-vision.rdz"
		    
		  Case "uvf", "uvvf", "uvd", "uvvd"
		    Return "application/vnd.dece.data"
		    
		  Case "uvt", "uvvt"
		    Return "application/vnd.dece.ttml+xml"
		    
		  Case "uvx", "uvvx"
		    Return "application/vnd.dece.unspecified"
		    
		  Case "fe_launch"
		    Return "application/vnd.denovo.fcselayout-link"
		    
		  Case "dna"
		    Return "application/vnd.dna"
		    
		  Case "mlp"
		    Return "application/vnd.dolby.mlp"
		    
		  Case "dpg"
		    Return "application/vnd.dpgraph"
		    
		  Case "dfac"
		    Return "application/vnd.dreamfactory"
		    
		  Case "ait"
		    Return "application/vnd.dvb.ait"
		    
		  Case "svc"
		    Return "application/vnd.dvb.service"
		    
		  Case "geo"
		    Return "application/vnd.dynageo"
		    
		  Case "mag"
		    Return "application/vnd.ecowin.chart"
		    
		  Case "nml"
		    Return "application/vnd.enliven"
		    
		  Case "esf"
		    Return "application/vnd.epson.esf"
		    
		  Case "msf"
		    Return "application/vnd.epson.msf"
		    
		  Case "qam"
		    Return "application/vnd.epson.quickanime"
		    
		  Case "slt"
		    Return "application/vnd.epson.salt"
		    
		  Case "ssf"
		    Return "application/vnd.epson.ssf"
		    
		  Case "es3", "et3"
		    Return "application/vnd.eszigno3+xml"
		    
		  Case "ez2"
		    Return "application/vnd.ezpix-album"
		    
		  Case "ez3"
		    Return "application/vnd.ezpix-package"
		    
		  Case "fdf"
		    Return "application/vnd.fdf"
		    
		  Case "mseed"
		    Return "application/vnd.fdsn.mseed"
		    
		  Case "seed", "dataless"
		    Return "application/vnd.fdsn.seed"
		    
		  Case "gph"
		    Return "application/vnd.flographit"
		    
		  Case "ftc"
		    Return "application/vnd.fluxtime.clip"
		    
		  Case "fm", "frame", "maker", "book"
		    Return "application/vnd.framemaker"
		    
		  Case "fnc"
		    Return "application/vnd.frogans.fnc"
		    
		  Case "ltf"
		    Return "application/vnd.frogans.ltf"
		    
		  Case "fsc"
		    Return "application/vnd.fsc.weblaunch"
		    
		  Case "oas"
		    Return "application/vnd.fujitsu.oasys"
		    
		  Case "oa2"
		    Return "application/vnd.fujitsu.oasys2"
		    
		  Case "oa3"
		    Return "application/vnd.fujitsu.oasys3"
		    
		  Case "fg5"
		    Return "application/vnd.fujitsu.oasysgp"
		    
		  Case "bh2"
		    Return "application/vnd.fujitsu.oasysprs"
		    
		  Case "ddd"
		    Return "application/vnd.fujixerox.ddd"
		    
		  Case "xdw"
		    Return "application/vnd.fujixerox.docuworks"
		    
		  Case "xbd"
		    Return "application/vnd.fujixerox.docuworks.binder"
		    
		  Case "fzs"
		    Return "application/vnd.fuzzysheet"
		    
		  Case "txd"
		    Return "application/vnd.genomatix.tuxedo"
		    
		  Case "ggb"
		    Return "application/vnd.geogebra.file"
		    
		  Case "ggt"
		    Return "application/vnd.geogebra.tool"
		    
		  Case "gex", "gre"
		    Return "application/vnd.geometry-explorer"
		    
		  Case "gxt"
		    Return "application/vnd.geonext"
		    
		  Case "g2w"
		    Return "application/vnd.geoplan"
		    
		  Case "g3w"
		    Return "application/vnd.geospace"
		    
		  Case "gmx"
		    Return "application/vnd.gmx"
		    
		  Case "kml"
		    Return "application/vnd.google-earth.kml+xml"
		    
		  Case "kmz"
		    Return "application/vnd.google-earth.kmz"
		    
		  Case "gqf", "gqs"
		    Return "application/vnd.grafeq"
		    
		  Case "gac"
		    Return "application/vnd.groove-account"
		    
		  Case "ghf"
		    Return "application/vnd.groove-help"
		    
		  Case "gim"
		    Return "application/vnd.groove-identity-message"
		    
		  Case "grv"
		    Return "application/vnd.groove-injector"
		    
		  Case "gtm"
		    Return "application/vnd.groove-tool-message"
		    
		  Case "tpl"
		    Return "application/vnd.groove-tool-template"
		    
		  Case "vcg"
		    Return "application/vnd.groove-vcard"
		    
		  Case "hal"
		    Return "application/vnd.hal+xml"
		    
		  Case "zmm"
		    Return "application/vnd.handheld-entertainment+xml"
		    
		  Case "hbci"
		    Return "application/vnd.hbci"
		    
		  Case "les"
		    Return "application/vnd.hhe.lesson-player"
		    
		  Case "hpgl"
		    Return "application/vnd.hp-hpgl"
		    
		  Case "hpid"
		    Return "application/vnd.hp-hpid"
		    
		  Case "hps"
		    Return "application/vnd.hp-hps"
		    
		  Case "jlt"
		    Return "application/vnd.hp-jlyt"
		    
		  Case "pcl"
		    Return "application/vnd.hp-pcl"
		    
		  Case "pclxl"
		    Return "application/vnd.hp-pclxl"
		    
		  Case "sfd-hdstx"
		    Return "application/vnd.hydrostatix.sof-data"
		    
		  Case "x3d"
		    Return "application/vnd.hzn-3d-crossword"
		    
		  Case "mpy"
		    Return "application/vnd.ibm.minipay"
		    
		  Case "afp", "listafp", "list3820"
		    Return "application/vnd.ibm.modcap"
		    
		  Case "irm"
		    Return "application/vnd.ibm.rights-management"
		    
		  Case "sc"
		    Return "application/vnd.ibm.secure-container"
		    
		  Case "icc", "icm"
		    Return "application/vnd.iccprofile"
		    
		  Case "igl"
		    Return "application/vnd.igloader"
		    
		  Case "ivp"
		    Return "application/vnd.immervision-ivp"
		    
		  Case "ivu"
		    Return "application/vnd.immervision-ivu"
		    
		  Case "igm"
		    Return "application/vnd.insors.igm"
		    
		  Case "xpw", "xpx"
		    Return "application/vnd.intercon.formnet"
		    
		  Case "i2g"
		    Return "application/vnd.intergeo"
		    
		  Case "qbo"
		    Return "application/vnd.intu.qbo"
		    
		  Case "qfx"
		    Return "application/vnd.intu.qfx"
		    
		  Case "rcprofile"
		    Return "application/vnd.ipunplugged.rcprofile"
		    
		  Case "irp"
		    Return "application/vnd.irepository.package+xml"
		    
		  Case "xpr"
		    Return "application/vnd.is-xpr"
		    
		  Case "fcs"
		    Return "application/vnd.isac.fcs"
		    
		  Case "jam"
		    Return "application/vnd.jam"
		    
		  Case "rms"
		    Return "application/vnd.jcp.javame.midlet-rms"
		    
		  Case "jisp"
		    Return "application/vnd.jisp"
		    
		  Case "joda"
		    Return "application/vnd.joost.joda-archive"
		    
		  Case "ktz", "ktr"
		    Return "application/vnd.kahootz"
		    
		  Case "karbon"
		    Return "application/vnd.kde.karbon"
		    
		  Case "chrt"
		    Return "application/vnd.kde.kchart"
		    
		  Case "kfo"
		    Return "application/vnd.kde.kformula"
		    
		  Case "flw"
		    Return "application/vnd.kde.kivio"
		    
		  Case "kon"
		    Return "application/vnd.kde.kontour"
		    
		  Case "kpr", "kpt"
		    Return "application/vnd.kde.kpresenter"
		    
		  Case "ksp"
		    Return "application/vnd.kde.kspread"
		    
		  Case "kwd", "kwt"
		    Return "application/vnd.kde.kword"
		    
		  Case "htke"
		    Return "application/vnd.kenameaapp"
		    
		  Case "kia"
		    Return "application/vnd.kidspiration"
		    
		  Case "kne", "knp"
		    Return "application/vnd.kinar"
		    
		  Case "skp", "skd", "skt", "skm"
		    Return "application/vnd.koan"
		    
		  Case "sse"
		    Return "application/vnd.kodak-descriptor"
		    
		  Case "lasxml"
		    Return "application/vnd.las.las+xml"
		    
		  Case "lbd"
		    Return "application/vnd.llamagraphics.life-balance.desktop"
		    
		  Case "lbe"
		    Return "application/vnd.llamagraphics.life-balance.exchange+xml"
		    
		  Case "123"
		    Return "application/vnd.lotus-1-2-3"
		    
		  Case "apr"
		    Return "application/vnd.lotus-approach"
		    
		  Case "pre"
		    Return "application/vnd.lotus-freelance"
		    
		  Case "nsf"
		    Return "application/vnd.lotus-notes"
		    
		  Case "org"
		    Return "application/vnd.lotus-organizer"
		    
		  Case "scm"
		    Return "application/vnd.lotus-screencam"
		    
		  Case "lwp"
		    Return "application/vnd.lotus-wordpro"
		    
		  Case "portpkg"
		    Return "application/vnd.macports.portpkg"
		    
		  Case "mcd"
		    Return "application/vnd.mcd"
		    
		  Case "mc1"
		    Return "application/vnd.medcalcdata"
		    
		  Case "cdkey"
		    Return "application/vnd.mediastation.cdkey"
		    
		  Case "mwf"
		    Return "application/vnd.mfer"
		    
		  Case "mfm"
		    Return "application/vnd.mfmp"
		    
		  Case "flo"
		    Return "application/vnd.micrografx.flo"
		    
		  Case "igx"
		    Return "application/vnd.micrografx.igx"
		    
		  Case "mif"
		    Return "application/vnd.mif"
		    
		  Case "daf"
		    Return "application/vnd.mobius.daf"
		    
		  Case "dis"
		    Return "application/vnd.mobius.dis"
		    
		  Case "mbk"
		    Return "application/vnd.mobius.mbk"
		    
		  Case "mqy"
		    Return "application/vnd.mobius.mqy"
		    
		  Case "msl"
		    Return "application/vnd.mobius.msl"
		    
		  Case "plc"
		    Return "application/vnd.mobius.plc"
		    
		  Case "txf"
		    Return "application/vnd.mobius.txf"
		    
		  Case "mpn"
		    Return "application/vnd.mophun.application"
		    
		  Case "mpc"
		    Return "application/vnd.mophun.certificate"
		    
		  Case "xul"
		    Return "application/vnd.mozilla.xul+xml"
		    
		  Case "cil"
		    Return "application/vnd.ms-artgalry"
		    
		  Case "cab"
		    Return "application/vnd.ms-cab-compressed"
		    
		  Case "xls", "xlm", "xla", "xlc", "xlt", "xlw"
		    Return "application/vnd.ms-excel"
		    
		  Case "xlam"
		    Return "application/vnd.ms-excel.addin.macroenabled.12"
		    
		  Case "xlsb"
		    Return "application/vnd.ms-excel.sheet.binary.macroenabled.12"
		    
		  Case "xlsm"
		    Return "application/vnd.ms-excel.sheet.macroenabled.12"
		    
		  Case "xltm"
		    Return "application/vnd.ms-excel.template.macroenabled.12"
		    
		  Case "eot"
		    Return "application/vnd.ms-fontobject"
		    
		  Case "chm"
		    Return "application/vnd.ms-htmlhelp"
		    
		  Case "ims"
		    Return "application/vnd.ms-ims"
		    
		  Case "lrm"
		    Return "application/vnd.ms-lrm"
		    
		  Case "thmx"
		    Return "application/vnd.ms-officetheme"
		    
		  Case "cat"
		    Return "application/vnd.ms-pki.seccat"
		    
		  Case "stl"
		    Return "application/vnd.ms-pki.stl"
		    
		  Case "ppt", "pps", "pot"
		    Return "application/vnd.ms-powerpoint"
		    
		  Case "ppam"
		    Return "application/vnd.ms-powerpoint.addin.macroenabled.12"
		    
		  Case "pptm"
		    Return "application/vnd.ms-powerpoint.presentation.macroenabled.12"
		    
		  Case "sldm"
		    Return "application/vnd.ms-powerpoint.slide.macroenabled.12"
		    
		  Case "ppsm"
		    Return "application/vnd.ms-powerpoint.slideshow.macroenabled.12"
		    
		  Case "potm"
		    Return "application/vnd.ms-powerpoint.template.macroenabled.12"
		    
		  Case "mpp", "mpt"
		    Return "application/vnd.ms-project"
		    
		  Case "docm"
		    Return "application/vnd.ms-word.document.macroenabled.12"
		    
		  Case "dotm"
		    Return "application/vnd.ms-word.template.macroenabled.12"
		    
		  Case "wps", "wks", "wcm", "wdb"
		    Return "application/vnd.ms-works"
		    
		  Case "wpl"
		    Return "application/vnd.ms-wpl"
		    
		  Case "xps"
		    Return "application/vnd.ms-xpsdocument"
		    
		  Case "mseq"
		    Return "application/vnd.mseq"
		    
		  Case "mus"
		    Return "application/vnd.musician"
		    
		  Case "msty"
		    Return "application/vnd.muvee.style"
		    
		  Case "nlu"
		    Return "application/vnd.neurolanguage.nlu"
		    
		  Case "nnd"
		    Return "application/vnd.noblenet-directory"
		    
		  Case "nns"
		    Return "application/vnd.noblenet-sealer"
		    
		  Case "nnw"
		    Return "application/vnd.noblenet-web"
		    
		  Case "ngdat"
		    Return "application/vnd.nokia.n-gage.data"
		    
		  Case "n-gage"
		    Return "application/vnd.nokia.n-gage.symbian.install"
		    
		  Case "rpst"
		    Return "application/vnd.nokia.radio-preset"
		    
		  Case "rpss"
		    Return "application/vnd.nokia.radio-presets"
		    
		  Case "edm"
		    Return "application/vnd.novadigm.edm"
		    
		  Case "edx"
		    Return "application/vnd.novadigm.edx"
		    
		  Case "ext"
		    Return "application/vnd.novadigm.ext"
		    
		  Case "odc"
		    Return "application/vnd.oasis.opendocument.chart"
		    
		  Case "otc"
		    Return "application/vnd.oasis.opendocument.chart-template"
		    
		  Case "odb"
		    Return "application/vnd.oasis.opendocument.database"
		    
		  Case "odf"
		    Return "application/vnd.oasis.opendocument.formula"
		    
		  Case "odft"
		    Return "application/vnd.oasis.opendocument.formula-template"
		    
		  Case "odg"
		    Return "application/vnd.oasis.opendocument.graphics"
		    
		  Case "otg"
		    Return "application/vnd.oasis.opendocument.graphics-template"
		    
		  Case "odi"
		    Return "application/vnd.oasis.opendocument.image"
		    
		  Case "oti"
		    Return "application/vnd.oasis.opendocument.image-template"
		    
		  Case "odp"
		    Return "application/vnd.oasis.opendocument.presentation"
		    
		  Case "otp"
		    Return "application/vnd.oasis.opendocument.presentation-template"
		    
		  Case "ods"
		    Return "application/vnd.oasis.opendocument.spreadsheet"
		    
		  Case "ots"
		    Return "application/vnd.oasis.opendocument.spreadsheet-template"
		    
		  Case "odt"
		    Return "application/vnd.oasis.opendocument.text"
		    
		  Case "odm"
		    Return "application/vnd.oasis.opendocument.text-master"
		    
		  Case "ott"
		    Return "application/vnd.oasis.opendocument.text-template"
		    
		  Case "oth"
		    Return "application/vnd.oasis.opendocument.text-web"
		    
		  Case "xo"
		    Return "application/vnd.olpc-sugar"
		    
		  Case "dd2"
		    Return "application/vnd.oma.dd2+xml"
		    
		  Case "oxt"
		    Return "application/vnd.openofficeorg.extension"
		    
		  Case "pptx"
		    Return "application/vnd.openxmlformats-officedocument.presentationml.presentation"
		    
		  Case "sldx"
		    Return "application/vnd.openxmlformats-officedocument.presentationml.slide"
		    
		  Case "ppsx"
		    Return "application/vnd.openxmlformats-officedocument.presentationml.slideshow"
		    
		  Case "potx"
		    Return "application/vnd.openxmlformats-officedocument.presentationml.template"
		    
		  Case "xlsx"
		    Return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
		    
		  Case "xltx"
		    Return "application/vnd.openxmlformats-officedocument.spreadsheetml.template"
		    
		  Case "docx"
		    Return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
		    
		  Case "dotx"
		    Return "application/vnd.openxmlformats-officedocument.wordprocessingml.template"
		    
		  Case "mgp"
		    Return "application/vnd.osgeo.mapguide.package"
		    
		  Case "dp"
		    Return "application/vnd.osgi.dp"
		    
		  Case "pdb", "pqa", "oprc"
		    Return "application/vnd.palm"
		    
		  Case "paw"
		    Return "application/vnd.pawaafile"
		    
		  Case "str"
		    Return "application/vnd.pg.format"
		    
		  Case "ei6"
		    Return "application/vnd.pg.osasli"
		    
		  Case "efif"
		    Return "application/vnd.picsel"
		    
		  Case "wg"
		    Return "application/vnd.pmi.widget"
		    
		  Case "plf"
		    Return "application/vnd.pocketlearn"
		    
		  Case "pbd"
		    Return "application/vnd.powerbuilder6"
		    
		  Case "box"
		    Return "application/vnd.previewsystems.box"
		    
		  Case "mgz"
		    Return "application/vnd.proteus.magazine"
		    
		  Case "qps"
		    Return "application/vnd.publishare-delta-tree"
		    
		  Case "ptid"
		    Return "application/vnd.pvi.ptid1"
		    
		  Case "qxd", "qxt", "qwd", "qwt", "qxl", "qxb"
		    Return "application/vnd.quark.quarkxpress"
		    
		  Case "bed"
		    Return "application/vnd.realvnc.bed"
		    
		  Case "mxl"
		    Return "application/vnd.recordare.musicxml"
		    
		  Case "musicxml"
		    Return "application/vnd.recordare.musicxml+xml"
		    
		  Case "cryptonote"
		    Return "application/vnd.rig.cryptonote"
		    
		  Case "cod"
		    Return "application/vnd.rim.cod"
		    
		  Case "rm"
		    Return "application/vnd.rn-realmedia"
		    
		  Case "link66"
		    Return "application/vnd.route66.link66+xml"
		    
		  Case "st"
		    Return "application/vnd.sailingtracker.track"
		    
		  Case "see"
		    Return "application/vnd.seemail"
		    
		  Case "sema"
		    Return "application/vnd.sema"
		    
		  Case "semd"
		    Return "application/vnd.semd"
		    
		  Case "semf"
		    Return "application/vnd.semf"
		    
		  Case "ifm"
		    Return "application/vnd.shana.informed.formdata"
		    
		  Case "itp"
		    Return "application/vnd.shana.informed.formtemplate"
		    
		  Case "iif"
		    Return "application/vnd.shana.informed.interchange"
		    
		  Case "ipk"
		    Return "application/vnd.shana.informed.package"
		    
		  Case "twd", "twds"
		    Return "application/vnd.simtech-mindmapper"
		    
		  Case "mmf"
		    Return "application/vnd.smaf"
		    
		  Case "teacher"
		    Return "application/vnd.smart.teacher"
		    
		  Case "sdkm", "sdkd"
		    Return "application/vnd.solent.sdkm+xml"
		    
		  Case "dxp"
		    Return "application/vnd.spotfire.dxp"
		    
		  Case "sfs"
		    Return "application/vnd.spotfire.sfs"
		    
		  Case "sdc"
		    Return "application/vnd.stardivision.calc"
		    
		  Case "sda"
		    Return "application/vnd.stardivision.draw"
		    
		  Case "sdd"
		    Return "application/vnd.stardivision.impress"
		    
		  Case "smf"
		    Return "application/vnd.stardivision.math"
		    
		  Case "sdw", "vor"
		    Return "application/vnd.stardivision.writer"
		    
		  Case "sgl"
		    Return "application/vnd.stardivision.writer-global"
		    
		  Case "sm"
		    Return "application/vnd.stepmania.stepchart"
		    
		  Case "sxc"
		    Return "application/vnd.sun.xml.calc"
		    
		  Case "stc"
		    Return "application/vnd.sun.xml.calc.template"
		    
		  Case "sxd"
		    Return "application/vnd.sun.xml.draw"
		    
		  Case "std"
		    Return "application/vnd.sun.xml.draw.template"
		    
		  Case "sxi"
		    Return "application/vnd.sun.xml.impress"
		    
		  Case "sti"
		    Return "application/vnd.sun.xml.impress.template"
		    
		  Case "sxm"
		    Return "application/vnd.sun.xml.math"
		    
		  Case "sxw"
		    Return "application/vnd.sun.xml.writer"
		    
		  Case "sxg"
		    Return "application/vnd.sun.xml.writer.global"
		    
		  Case "stw"
		    Return "application/vnd.sun.xml.writer.template"
		    
		  Case "sus", "susp"
		    Return "application/vnd.sus-calendar"
		    
		  Case "svd"
		    Return "application/vnd.svd"
		    
		  Case "sis", "sisx"
		    Return "application/vnd.symbian.install"
		    
		  Case "xsm"
		    Return "application/vnd.syncml+xml"
		    
		  Case "bdm"
		    Return "application/vnd.syncml.dm+wbxml"
		    
		  Case "xdm"
		    Return "application/vnd.syncml.dm+xml"
		    
		  Case "tao"
		    Return "application/vnd.tao.intent-module-archive"
		    
		  Case "tmo"
		    Return "application/vnd.tmobile-livetv"
		    
		  Case "tpt"
		    Return "application/vnd.trid.tpt"
		    
		  Case "mxs"
		    Return "application/vnd.triscape.mxs"
		    
		  Case "tra"
		    Return "application/vnd.trueapp"
		    
		  Case "ufd", "ufdl"
		    Return "application/vnd.ufdl"
		    
		  Case "utz"
		    Return "application/vnd.uiq.theme"
		    
		  Case "umj"
		    Return "application/vnd.umajin"
		    
		  Case "unityweb"
		    Return "application/vnd.unity"
		    
		  Case "uoml"
		    Return "application/vnd.uoml+xml"
		    
		  Case "vcx"
		    Return "application/vnd.vcx"
		    
		  Case "vsd", "vst", "vss", "vsw"
		    Return "application/vnd.visio"
		    
		  Case "vis"
		    Return "application/vnd.visionary"
		    
		  Case "vsf"
		    Return "application/vnd.vsf"
		    
		  Case "wbxml"
		    Return "application/vnd.wap.wbxml"
		    
		  Case "wmlc"
		    Return "application/vnd.wap.wmlc"
		    
		  Case "wmlsc"
		    Return "application/vnd.wap.wmlscriptc"
		    
		  Case "wtb"
		    Return "application/vnd.webturbo"
		    
		  Case "nbp"
		    Return "application/vnd.wolfram.player"
		    
		  Case "wpd"
		    Return "application/vnd.wordperfect"
		    
		  Case "wqd"
		    Return "application/vnd.wqd"
		    
		  Case "stf"
		    Return "application/vnd.wt.stf"
		    
		  Case "xar"
		    Return "application/vnd.xara"
		    
		  Case "xfdl"
		    Return "application/vnd.xfdl"
		    
		  Case "hvd"
		    Return "application/vnd.yamaha.hv-dic"
		    
		  Case "hvs"
		    Return "application/vnd.yamaha.hv-script"
		    
		  Case "hvp"
		    Return "application/vnd.yamaha.hv-voice"
		    
		  Case "osf"
		    Return "application/vnd.yamaha.openscoreformat"
		    
		  Case "osfpvg"
		    Return "application/vnd.yamaha.openscoreformat.osfpvg+xml"
		    
		  Case "saf"
		    Return "application/vnd.yamaha.smaf-audio"
		    
		  Case "spf"
		    Return "application/vnd.yamaha.smaf-phrase"
		    
		  Case "cmp"
		    Return "application/vnd.yellowriver-custom-menu"
		    
		  Case "zir", "zirz"
		    Return "application/vnd.zul"
		    
		  Case "zaz"
		    Return "application/vnd.zzazz.deck+xml"
		    
		  Case "vxml"
		    Return "application/voicexml+xml"
		    
		  Case "wgt"
		    Return "application/widget"
		    
		  Case "hlp"
		    Return "application/winhlp"
		    
		  Case "wsdl"
		    Return "application/wsdl+xml"
		    
		  Case "wspolicy"
		    Return "application/wspolicy+xml"
		    
		  Case "7z"
		    Return "application/x-7z-compressed"
		    
		  Case "abw"
		    Return "application/x-abiword"
		    
		  Case "ace"
		    Return "application/x-ace-compressed"
		    
		  Case "aab", "x32", "u32", "vox"
		    Return "application/x-authorware-bin"
		    
		  Case "aam"
		    Return "application/x-authorware-map"
		    
		  Case "aas"
		    Return "application/x-authorware-seg"
		    
		  Case "bcpio"
		    Return "application/x-bcpio"
		    
		  Case "torrent"
		    Return "application/x-bittorrent"
		    
		  Case "bz"
		    Return "application/x-bzip"
		    
		  Case "bz2", "boz"
		    Return "application/x-bzip2"
		    
		  Case "vcd"
		    Return "application/x-cdlink"
		    
		  Case "chat"
		    Return "application/x-chat"
		    
		  Case "pgn"
		    Return "application/x-chess-pgn"
		    
		  Case "cpio"
		    Return "application/x-cpio"
		    
		  Case "csh"
		    Return "application/x-csh"
		    
		  Case "deb", "udeb"
		    Return "application/x-debian-package"
		    
		  Case "dir", "dcr", "dxr", "cst", "cct", "cxt", "w3d", "fgd", "swa"
		    Return "application/x-director"
		    
		  Case "wad"
		    Return "application/x-doom"
		    
		  Case "ncx"
		    Return "application/x-dtbncx+xml"
		    
		  Case "dtb"
		    Return "application/x-dtbook+xml"
		    
		  Case "res"
		    Return "application/x-dtbresource+xml"
		    
		  Case "dvi"
		    Return "application/x-dvi"
		    
		  Case "bdf"
		    Return "application/x-font-bdf"
		    
		  Case "gsf"
		    Return "application/x-font-ghostscript"
		    
		  Case "psf"
		    Return "application/x-font-linux-psf"
		    
		  Case "otf"
		    Return "application/x-font-otf"
		    
		  Case "pcf"
		    Return "application/x-font-pcf"
		    
		  Case "snf"
		    Return "application/x-font-snf"
		    
		  Case "ttf", "ttc"
		    Return "application/x-font-ttf"
		    
		  Case "pfa", "pfb", "pfm", "afm"
		    Return "application/x-font-type1"
		    
		  Case "woff"
		    Return "application/x-font-woff"
		    
		  Case "spl"
		    Return "application/x-futuresplash"
		    
		  Case "gnumeric"
		    Return "application/x-gnumeric"
		    
		  Case "gtar"
		    Return "application/x-gtar"
		    
		  Case "hdf"
		    Return "application/x-hdf"
		    
		  Case "jnlp"
		    Return "application/x-java-jnlp-file"
		    
		  Case "latex"
		    Return "application/x-latex"
		    
		  Case "prc", "mobi"
		    Return "application/x-mobipocket-ebook"
		    
		  Case "m3u8"
		    Return "application/x-mpegurl"
		    
		  Case "application"
		    Return "application/x-ms-application"
		    
		  Case "wmd"
		    Return "application/x-ms-wmd"
		    
		  Case "wmz"
		    Return "application/x-ms-wmz"
		    
		  Case "xbap"
		    Return "application/x-ms-xbap"
		    
		  Case "mdb"
		    Return "application/x-msaccess"
		    
		  Case "obd"
		    Return "application/x-msbinder"
		    
		  Case "crd"
		    Return "application/x-mscardfile"
		    
		  Case "clp"
		    Return "application/x-msclip"
		    
		  Case "exe", "dll", "com", "bat", "msi"
		    Return "application/x-msdownload"
		    
		  Case "mvb", "m13", "m14"
		    Return "application/x-msmediaview"
		    
		  Case "wmf"
		    Return "application/x-msmetafile"
		    
		  Case "mny"
		    Return "application/x-msmoney"
		    
		  Case "pub"
		    Return "application/x-mspublisher"
		    
		  Case "scd"
		    Return "application/x-msschedule"
		    
		  Case "trm"
		    Return "application/x-msterminal"
		    
		  Case "wri"
		    Return "application/x-mswrite"
		    
		  Case "nc", "cdf"
		    Return "application/x-netcdf"
		    
		  Case "p12", "pfx"
		    Return "application/x-pkcs12"
		    
		  Case "p7b", "spc"
		    Return "application/x-pkcs7-certificates"
		    
		  Case "p7r"
		    Return "application/x-pkcs7-certreqresp"
		    
		  Case "rar"
		    Return "application/x-rar-compressed"
		    
		  Case "sh"
		    Return "application/x-sh"
		    
		  Case "shar"
		    Return "application/x-shar"
		    
		  Case "swf"
		    Return "application/x-shockwave-flash"
		    
		  Case "xap"
		    Return "application/x-silverlight-app"
		    
		  Case "sit"
		    Return "application/x-stuffit"
		    
		  Case "sitx"
		    Return "application/x-stuffitx"
		    
		  Case "sv4cpio"
		    Return "application/x-sv4cpio"
		    
		  Case "sv4crc"
		    Return "application/x-sv4crc"
		    
		  Case "tar"
		    Return "application/x-tar"
		    
		  Case "tcl"
		    Return "application/x-tcl"
		    
		  Case "tex"
		    Return "application/x-tex"
		    
		  Case "tfm"
		    Return "application/x-tex-tfm"
		    
		  Case "texinfo", "texi"
		    Return "application/x-texinfo"
		    
		  Case "ustar"
		    Return "application/x-ustar"
		    
		  Case "src"
		    Return "application/x-wais-source"
		    
		  Case "der", "crt"
		    Return "application/x-x509-ca-cert"
		    
		  Case "fig"
		    Return "application/x-xfig"
		    
		  Case "xpi"
		    Return "application/x-xpinstall"
		    
		  Case "xdf"
		    Return "application/xcap-diff+xml"
		    
		  Case "xenc"
		    Return "application/xenc+xml"
		    
		  Case "xhtml", "xht"
		    Return "application/xhtml+xml"
		    
		  Case "xml", "xsl"
		    Return "application/xml"
		    
		  Case "dtd"
		    Return "application/xml-dtd"
		    
		  Case "xop"
		    Return "application/xop+xml"
		    
		  Case "xslt"
		    Return "application/xslt+xml"
		    
		  Case "xspf"
		    Return "application/xspf+xml"
		    
		  Case "mxml", "xhvml", "xvml", "xvm"
		    Return "application/xv+xml"
		    
		  Case "yang"
		    Return "application/yang"
		    
		  Case "yin"
		    Return "application/yin+xml"
		    
		  Case "zip"
		    Return "application/zip"
		    
		  Case "adp"
		    Return "audio/adpcm"
		    
		  Case "au", "snd"
		    Return "audio/basic"
		    
		  Case "mid", "midi", "kar", "rmi"
		    Return "audio/midi"
		    
		  Case "mp4a"
		    Return "audio/mp4"
		    
		  Case "m4a", "m4p"
		    Return "audio/mp4a-latm"
		    
		  Case "mpga", "mp2", "mp2a", "mp3", "m2a", "m3a"
		    Return "audio/mpeg"
		    
		  Case "oga", "ogg", "spx"
		    Return "audio/ogg"
		    
		  Case "uva", "uvva"
		    Return "audio/vnd.dece.audio"
		    
		  Case "eol"
		    Return "audio/vnd.digital-winds"
		    
		  Case "dra"
		    Return "audio/vnd.dra"
		    
		  Case "dts"
		    Return "audio/vnd.dts"
		    
		  Case "dtshd"
		    Return "audio/vnd.dts.hd"
		    
		  Case "lvp"
		    Return "audio/vnd.lucent.voice"
		    
		  Case "pya"
		    Return "audio/vnd.ms-playready.media.pya"
		    
		  Case "ecelp4800"
		    Return "audio/vnd.nuera.ecelp4800"
		    
		  Case "ecelp7470"
		    Return "audio/vnd.nuera.ecelp7470"
		    
		  Case "ecelp9600"
		    Return "audio/vnd.nuera.ecelp9600"
		    
		  Case "rip"
		    Return "audio/vnd.rip"
		    
		  Case "weba"
		    Return "audio/webm"
		    
		  Case "aac"
		    Return "audio/x-aac"
		    
		  Case "aif", "aiff", "aifc"
		    Return "audio/x-aiff"
		    
		  Case "m3u"
		    Return "audio/x-mpegurl"
		    
		  Case "wax"
		    Return "audio/x-ms-wax"
		    
		  Case "wma"
		    Return "audio/x-ms-wma"
		    
		  Case "ram", "ra"
		    Return "audio/x-pn-realaudio"
		    
		  Case "rmp"
		    Return "audio/x-pn-realaudio-plugin"
		    
		  Case "wav"
		    Return "audio/x-wav"
		    
		  Case "cdx"
		    Return "chemical/x-cdx"
		    
		  Case "cif"
		    Return "chemical/x-cif"
		    
		  Case "cmdf"
		    Return "chemical/x-cmdf"
		    
		  Case "cml"
		    Return "chemical/x-cml"
		    
		  Case "csml"
		    Return "chemical/x-csml"
		    
		  Case "xyz"
		    Return "chemical/x-xyz"
		    
		  Case "bmp"
		    Return "image/bmp"
		    
		  Case "cgm"
		    Return "image/cgm"
		    
		  Case "g3"
		    Return "image/g3fax"
		    
		  Case "gif"
		    Return "image/gif"
		    
		  Case "ief"
		    Return "image/ief"
		    
		  Case "jp2"
		    Return "image/jp2"
		    
		  Case "jpeg", "jpg", "jpe"
		    Return "image/jpeg"
		    
		  Case "ktx"
		    Return "image/ktx"
		    
		  Case "pict", "pic", "pct"
		    Return "image/pict"
		    
		  Case "png"
		    Return "image/png"
		    
		  Case "btif"
		    Return "image/prs.btif"
		    
		  Case "svg", "svgz"
		    Return "image/svg+xml"
		    
		  Case "tiff", "tif"
		    Return "image/tiff"
		    
		  Case "psd"
		    Return "image/vnd.adobe.photoshop"
		    
		  Case "uvi", "uvvi", "uvg", "uvvg"
		    Return "image/vnd.dece.graphic"
		    
		  Case "sub"
		    Return "image/vnd.dvb.subtitle"
		    
		  Case "djvu", "djv"
		    Return "image/vnd.djvu"
		    
		  Case "dwg"
		    Return "image/vnd.dwg"
		    
		  Case "dxf"
		    Return "image/vnd.dxf"
		    
		  Case "fbs"
		    Return "image/vnd.fastbidsheet"
		    
		  Case "fpx"
		    Return "image/vnd.fpx"
		    
		  Case "fst"
		    Return "image/vnd.fst"
		    
		  Case "mmr"
		    Return "image/vnd.fujixerox.edmics-mmr"
		    
		  Case "rlc"
		    Return "image/vnd.fujixerox.edmics-rlc"
		    
		  Case "mdi"
		    Return "image/vnd.ms-modi"
		    
		  Case "npx"
		    Return "image/vnd.net-fpx"
		    
		  Case "wbmp"
		    Return "image/vnd.wap.wbmp"
		    
		  Case "xif"
		    Return "image/vnd.xiff"
		    
		  Case "webp"
		    Return "image/webp"
		    
		  Case "ras"
		    Return "image/x-cmu-raster"
		    
		  Case "cmx"
		    Return "image/x-cmx"
		    
		  Case "fh", "fhc", "fh4", "fh5", "fh7"
		    Return "image/x-freehand"
		    
		  Case "ico"
		    Return "image/x-icon"
		    
		  Case "pntg", "pnt", "mac"
		    Return "image/x-macpaint"
		    
		  Case "pcx"
		    Return "image/x-pcx"
		    
		  Case "pic", "pct"
		    Return "image/x-pict"
		    
		  Case "pnm"
		    Return "image/x-portable-anymap"
		    
		  Case "pbm"
		    Return "image/x-portable-bitmap"
		    
		  Case "pgm"
		    Return "image/x-portable-graymap"
		    
		  Case "ppm"
		    Return "image/x-portable-pixmap"
		    
		  Case "qtif", "qti"
		    Return "image/x-quicktime"
		    
		  Case "rgb"
		    Return "image/x-rgb"
		    
		  Case "xbm"
		    Return "image/x-xbitmap"
		    
		  Case "xpm"
		    Return "image/x-xpixmap"
		    
		  Case "xwd"
		    Return "image/x-xwindowdump"
		    
		  Case "eml", "mime"
		    Return "message/rfc822"
		    
		  Case "igs", "iges"
		    Return "model/iges"
		    
		  Case "msh", "mesh", "silo"
		    Return "model/mesh"
		    
		  Case "dae"
		    Return "model/vnd.collada+xml"
		    
		  Case "dwf"
		    Return "model/vnd.dwf"
		    
		  Case "gdl"
		    Return "model/vnd.gdl"
		    
		  Case "gtw"
		    Return "model/vnd.gtw"
		    
		  Case "mts"
		    Return "model/vnd.mts"
		    
		  Case "vtu"
		    Return "model/vnd.vtu"
		    
		  Case "wrl", "vrml"
		    Return "model/vrml"
		    
		  Case "manifest"
		    Return "text/cache-manifest"
		    
		  Case "ics", "ifb"
		    Return "text/calendar"
		    
		  Case "css"
		    Return "text/css"
		    
		  Case "csv"
		    Return "text/csv"
		    
		  Case "html", "htm"
		    Return "text/html"
		    
		  Case "n3"
		    Return "text/n3"
		    
		  Case "txt", "text", "conf", "def", "list", "log", "in"
		    Return "text/plain"
		    
		  Case "dsc"
		    Return "text/prs.lines.tag"
		    
		  Case "rtx"
		    Return "text/richtext"
		    
		  Case "sgml", "sgm"
		    Return "text/sgml"
		    
		  Case "tsv"
		    Return "text/tab-separated-values"
		    
		  Case "t", "tr", "roff", "man", "me", "ms"
		    Return "text/troff"
		    
		  Case "ttl"
		    Return "text/turtle"
		    
		  Case "uri", "uris", "urls"
		    Return "text/uri-list"
		    
		  Case "curl"
		    Return "text/vnd.curl"
		    
		  Case "dcurl"
		    Return "text/vnd.curl.dcurl"
		    
		  Case "scurl"
		    Return "text/vnd.curl.scurl"
		    
		  Case "mcurl"
		    Return "text/vnd.curl.mcurl"
		    
		  Case "fly"
		    Return "text/vnd.fly"
		    
		  Case "flx"
		    Return "text/vnd.fmi.flexstor"
		    
		  Case "gv"
		    Return "text/vnd.graphviz"
		    
		  Case "3dml"
		    Return "text/vnd.in3d.3dml"
		    
		  Case "spot"
		    Return "text/vnd.in3d.spot"
		    
		  Case "jad"
		    Return "text/vnd.sun.j2me.app-descriptor"
		    
		  Case "wml"
		    Return "text/vnd.wap.wml"
		    
		  Case "wmls"
		    Return "text/vnd.wap.wmlscript"
		    
		  Case "s", "asm"
		    Return "text/x-asm"
		    
		  Case "c", "cc", "cxx", "cpp", "h", "hh", "dic"
		    Return "text/x-c"
		    
		  Case "f", "for", "f77", "f90"
		    Return "text/x-fortran"
		    
		  Case "p", "pas"
		    Return "text/x-pascal"
		    
		  Case "java"
		    Return "text/x-java-source"
		    
		  Case "etx"
		    Return "text/x-setext"
		    
		  Case "uu"
		    Return "text/x-uuencode"
		    
		  Case "vcs"
		    Return "text/x-vcalendar"
		    
		  Case "vcf"
		    Return "text/x-vcard"
		    
		  Case "3gp"
		    Return "video/3gpp"
		    
		  Case "3g2"
		    Return "video/3gpp2"
		    
		  Case "h261"
		    Return "video/h261"
		    
		  Case "h263"
		    Return "video/h263"
		    
		  Case "h264"
		    Return "video/h264"
		    
		  Case "jpgv"
		    Return "video/jpeg"
		    
		  Case "jpm", "jpgm"
		    Return "video/jpm"
		    
		  Case "mj2", "mjp2"
		    Return "video/mj2"
		    
		  Case "ts"
		    Return "video/mp2t"
		    
		  Case "mp4", "mp4v", "mpg4", "m4v"
		    Return "video/mp4"
		    
		  Case "mpeg", "mpg", "mpe", "m1v", "m2v"
		    Return "video/mpeg"
		    
		  Case "ogv"
		    Return "video/ogg"
		    
		  Case "qt", "mov"
		    Return "video/quicktime"
		    
		  Case "uvh", "uvvh"
		    Return "video/vnd.dece.hd"
		    
		  Case "uvm", "uvvm"
		    Return "video/vnd.dece.mobile"
		    
		  Case "uvp", "uvvp"
		    Return "video/vnd.dece.pd"
		    
		  Case "uvs", "uvvs"
		    Return "video/vnd.dece.sd"
		    
		  Case "uvv", "uvvv"
		    Return "video/vnd.dece.video"
		    
		  Case "fvt"
		    Return "video/vnd.fvt"
		    
		  Case "mxu", "m4u"
		    Return "video/vnd.mpegurl"
		    
		  Case "pyv"
		    Return "video/vnd.ms-playready.media.pyv"
		    
		  Case "uvu", "uvvu"
		    Return "video/vnd.uvvu.mp4"
		    
		  Case "viv"
		    Return "video/vnd.vivo"
		    
		  Case "dv", "dif"
		    Return "video/x-dv"
		    
		  Case "webm"
		    Return "video/webm"
		    
		  Case "f4v"
		    Return "video/x-f4v"
		    
		  Case "fli"
		    Return "video/x-fli"
		    
		  Case "flv"
		    Return "video/x-flv"
		    
		  Case "m4v"
		    Return "video/x-m4v"
		    
		  Case "asf", "asx"
		    Return "video/x-ms-asf"
		    
		  Case "wm"
		    Return "video/x-ms-wm"
		    
		  Case "wmv"
		    Return "video/x-ms-wmv"
		    
		  Case "wmx"
		    Return "video/x-ms-wmx"
		    
		  Case "wvx"
		    Return "video/x-ms-wvx"
		    
		  Case "avi"
		    Return "video/x-msvideo"
		    
		  Case "movie"
		    Return "video/x-sgi-movie"
		    
		  Case "ice"
		    Return "x-conference/x-cooltalk"
		    
		  Else
		    ' This returns the default mime type
		    Return "application/octet-stream"
		    'Return "text/plain"
		    
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function RequestRescan(ResourceID As String, APIKey As String) As JSONItem
		  Dim js As New JSONItem
		  js.Value("resource") =ResourceID
		  js.Value("apikey") = APIKey
		  Return SendRequest(VT_Rescan_File, js)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function SendRequest(URL As String, Request As JSONItem, VTSock As HTTPSecureSocket = Nil, Timeout As Integer = 5) As JSONItem
		  If VTSock = Nil Then VTSock = New HTTPSecureSocket
		  VTSock.SetRequestHeader("User-Agent", "RB-VTAPI")
		  VTSock.Secure = True
		  VTSock.ConnectionType = VTSock.TLSv1
		  dim formData As New Dictionary
		  For Each Name As String In Request.Names
		    formData.Value(Name) = Request.Value(Name)
		  Next
		  VTSock.SetFormData(formData)
		  Dim s As String = VTSock.Post(URL, Timeout)
		  Dim js As JSONItem
		  Try
		    js = New JSONItem(s)
		    LastResponseCode = js.Value("response_code")
		    If js.HasName("verbose_msg") Then LastResponseVerbose = js.Value("verbose_msg")
		    Return js
		  Catch Err As JSONException
		    If VTSock.LastErrorCode = 0 Then
		      LastResponseCode = INVALID_RESPONSE
		      LastResponseVerbose = "VirusTotal.com responded with improperly formatted data."
		    Else
		      LastResponseCode = VTSock.LastErrorCode
		      LastResponseVerbose = SocketErrorMessage(VTSock)
		    End If
		    js = Nil
		  End Try
		  
		  Return js
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function SocketErrorMessage(Sender As SocketCore) As String
		  Dim err As String = "Socket error " + Str(Sender.LastErrorCode)
		  Select Case Sender.LastErrorCode
		  Case 102
		    err = err + ": Disconnected."
		  Case 100
		    err = err + ": Could not create a socket!"
		  Case 103
		    err = err + ": Connection timed out."
		  Case 105
		    err = err + ": That port number is already in use."
		  Case 106
		    err = err + ": Socket is not ready for that command."
		  Case 107
		    err = err + ": Could not bind to port."
		  Case 108
		    err = err + ": Out of memory."
		  Else
		    err = err + ": System error code."
		  End Select
		  
		  Return err
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function SubmitFile(File As FolderItem, APIKey As String) As JSONItem
		  'FIME
		  'Please note that this method doesn't actually work.
		  
		  Dim js As New JSONItem
		  Dim sock As New HTTPSecureSocket
		  
		  js.Value("file") = File.Name
		  js.Value("apikey") = APIKey
		  Dim upload As String = ConstructUpload(File)  '<-- doesn't work. help? https://www.virustotal.com/documentation/public-api/#scanning-files
		  sock.SetRequestContent(upload, "multipart/form-data, boundary=" + MIMEBoundary)
		  Return VTAPI.SendRequest(VT_Submit_File, js, sock, 0)
		  
		End Function
	#tag EndMethod


	#tag Note, Name = Copying
		Copyright ©2012 Andrew Lambert, All Rights Reserved.
		
		This program is free software: you can redistribute it and/or modify it under the terms of the GNU Affero General Public License as 
		published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
		
		This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of
		MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Affero General Public License for more details.
		
		You should have received a copy of the GNU Affero General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.
		
		---
		
		                    GNU AFFERO GENERAL PUBLIC LICENSE
		                       Version 3, 19 November 2007
		
		 Copyright (C) 2007 Free Software Foundation, Inc. <http://fsf.org/>
		 Everyone is permitted to copy and distribute verbatim copies
		 of this license document, but changing it is not allowed.
		
		                            Preamble
		
		  The GNU Affero General Public License is a free, copyleft license for
		software and other kinds of works, specifically designed to ensure
		cooperation with the community in the case of network server software.
		
		  The licenses for most software and other practical works are designed
		to take away your freedom to share and change the works.  By contrast,
		our General Public Licenses are intended to guarantee your freedom to
		share and change all versions of a program--to make sure it remains free
		software for all its users.
		
		  When we speak of free software, we are referring to freedom, not
		price.  Our General Public Licenses are designed to make sure that you
		have the freedom to distribute copies of free software (and charge for
		them if you wish), that you receive source code or can get it if you
		want it, that you can change the software or use pieces of it in new
		free programs, and that you know you can do these things.
		
		  Developers that use our General Public Licenses protect your rights
		with two steps: (1) assert copyright on the software, and (2) offer
		you this License which gives you legal permission to copy, distribute
		and/or modify the software.
		
		  A secondary benefit of defending all users' freedom is that
		improvements made in alternate versions of the program, if they
		receive widespread use, become available for other developers to
		incorporate.  Many developers of free software are heartened and
		encouraged by the resulting cooperation.  However, in the case of
		software used on network servers, this result may fail to come about.
		The GNU General Public License permits making a modified version and
		letting the public access it on a server without ever releasing its
		source code to the public.
		
		  The GNU Affero General Public License is designed specifically to
		ensure that, in such cases, the modified source code becomes available
		to the community.  It requires the operator of a network server to
		provide the source code of the modified version running there to the
		users of that server.  Therefore, public use of a modified version, on
		a publicly accessible server, gives the public access to the source
		code of the modified version.
		
		  An older license, called the Affero General Public License and
		published by Affero, was designed to accomplish similar goals.  This is
		a different license, not a version of the Affero GPL, but Affero has
		released a new version of the Affero GPL which permits relicensing under
		this license.
		
		  The precise terms and conditions for copying, distribution and
		modification follow.
		
		                       TERMS AND CONDITIONS
		
		  0. Definitions.
		
		  "This License" refers to version 3 of the GNU Affero General Public License.
		
		  "Copyright" also means copyright-like laws that apply to other kinds of
		works, such as semiconductor masks.
		
		  "The Program" refers to any copyrightable work licensed under this
		License.  Each licensee is addressed as "you".  "Licensees" and
		"recipients" may be individuals or organizations.
		
		  To "modify" a work means to copy from or adapt all or part of the work
		in a fashion requiring copyright permission, other than the making of an
		exact copy.  The resulting work is called a "modified version" of the
		earlier work or a work "based on" the earlier work.
		
		  A "covered work" means either the unmodified Program or a work based
		on the Program.
		
		  To "propagate" a work means to do anything with it that, without
		permission, would make you directly or secondarily liable for
		infringement under applicable copyright law, except executing it on a
		computer or modifying a private copy.  Propagation includes copying,
		distribution (with or without modification), making available to the
		public, and in some countries other activities as well.
		
		  To "convey" a work means any kind of propagation that enables other
		parties to make or receive copies.  Mere interaction with a user through
		a computer network, with no transfer of a copy, is not conveying.
		
		  An interactive user interface displays "Appropriate Legal Notices"
		to the extent that it includes a convenient and prominently visible
		feature that (1) displays an appropriate copyright notice, and (2)
		tells the user that there is no warranty for the work (except to the
		extent that warranties are provided), that licensees may convey the
		work under this License, and how to view a copy of this License.  If
		the interface presents a list of user commands or options, such as a
		menu, a prominent item in the list meets this criterion.
		
		  1. Source Code.
		
		  The "source code" for a work means the preferred form of the work
		for making modifications to it.  "Object code" means any non-source
		form of a work.
		
		  A "Standard Interface" means an interface that either is an official
		standard defined by a recognized standards body, or, in the case of
		interfaces specified for a particular programming language, one that
		is widely used among developers working in that language.
		
		  The "System Libraries" of an executable work include anything, other
		than the work as a whole, that (a) is included in the normal form of
		packaging a Major Component, but which is not part of that Major
		Component, and (b) serves only to enable use of the work with that
		Major Component, or to implement a Standard Interface for which an
		implementation is available to the public in source code form.  A
		"Major Component", in this context, means a major essential component
		(kernel, window system, and so on) of the specific operating system
		(if any) on which the executable work runs, or a compiler used to
		produce the work, or an object code interpreter used to run it.
		
		  The "Corresponding Source" for a work in object code form means all
		the source code needed to generate, install, and (for an executable
		work) run the object code and to modify the work, including scripts to
		control those activities.  However, it does not include the work's
		System Libraries, or general-purpose tools or generally available free
		programs which are used unmodified in performing those activities but
		which are not part of the work.  For example, Corresponding Source
		includes interface definition files associated with source files for
		the work, and the source code for shared libraries and dynamically
		linked subprograms that the work is specifically designed to require,
		such as by intimate data communication or control flow between those
		subprograms and other parts of the work.
		
		  The Corresponding Source need not include anything that users
		can regenerate automatically from other parts of the Corresponding
		Source.
		
		  The Corresponding Source for a work in source code form is that
		same work.
		
		  2. Basic Permissions.
		
		  All rights granted under this License are granted for the term of
		copyright on the Program, and are irrevocable provided the stated
		conditions are met.  This License explicitly affirms your unlimited
		permission to run the unmodified Program.  The output from running a
		covered work is covered by this License only if the output, given its
		content, constitutes a covered work.  This License acknowledges your
		rights of fair use or other equivalent, as provided by copyright law.
		
		  You may make, run and propagate covered works that you do not
		convey, without conditions so long as your license otherwise remains
		in force.  You may convey covered works to others for the sole purpose
		of having them make modifications exclusively for you, or provide you
		with facilities for running those works, provided that you comply with
		the terms of this License in conveying all material for which you do
		not control copyright.  Those thus making or running the covered works
		for you must do so exclusively on your behalf, under your direction
		and control, on terms that prohibit them from making any copies of
		your copyrighted material outside their relationship with you.
		
		  Conveying under any other circumstances is permitted solely under
		the conditions stated below.  Sublicensing is not allowed; section 10
		makes it unnecessary.
		
		  3. Protecting Users' Legal Rights From Anti-Circumvention Law.
		
		  No covered work shall be deemed part of an effective technological
		measure under any applicable law fulfilling obligations under article
		11 of the WIPO copyright treaty adopted on 20 December 1996, or
		similar laws prohibiting or restricting circumvention of such
		measures.
		
		  When you convey a covered work, you waive any legal power to forbid
		circumvention of technological measures to the extent such circumvention
		is effected by exercising rights under this License with respect to
		the covered work, and you disclaim any intention to limit operation or
		modification of the work as a means of enforcing, against the work's
		users, your or third parties' legal rights to forbid circumvention of
		technological measures.
		
		  4. Conveying Verbatim Copies.
		
		  You may convey verbatim copies of the Program's source code as you
		receive it, in any medium, provided that you conspicuously and
		appropriately publish on each copy an appropriate copyright notice;
		keep intact all notices stating that this License and any
		non-permissive terms added in accord with section 7 apply to the code;
		keep intact all notices of the absence of any warranty; and give all
		recipients a copy of this License along with the Program.
		
		  You may charge any price or no price for each copy that you convey,
		and you may offer support or warranty protection for a fee.
		
		  5. Conveying Modified Source Versions.
		
		  You may convey a work based on the Program, or the modifications to
		produce it from the Program, in the form of source code under the
		terms of section 4, provided that you also meet all of these conditions:
		
		    a) The work must carry prominent notices stating that you modified
		    it, and giving a relevant date.
		
		    b) The work must carry prominent notices stating that it is
		    released under this License and any conditions added under section
		    7.  This requirement modifies the requirement in section 4 to
		    "keep intact all notices".
		
		    c) You must license the entire work, as a whole, under this
		    License to anyone who comes into possession of a copy.  This
		    License will therefore apply, along with any applicable section 7
		    additional terms, to the whole of the work, and all its parts,
		    regardless of how they are packaged.  This License gives no
		    permission to license the work in any other way, but it does not
		    invalidate such permission if you have separately received it.
		
		    d) If the work has interactive user interfaces, each must display
		    Appropriate Legal Notices; however, if the Program has interactive
		    interfaces that do not display Appropriate Legal Notices, your
		    work need not make them do so.
		
		  A compilation of a covered work with other separate and independent
		works, which are not by their nature extensions of the covered work,
		and which are not combined with it such as to form a larger program,
		in or on a volume of a storage or distribution medium, is called an
		"aggregate" if the compilation and its resulting copyright are not
		used to limit the access or legal rights of the compilation's users
		beyond what the individual works permit.  Inclusion of a covered work
		in an aggregate does not cause this License to apply to the other
		parts of the aggregate.
		
		  6. Conveying Non-Source Forms.
		
		  You may convey a covered work in object code form under the terms
		of sections 4 and 5, provided that you also convey the
		machine-readable Corresponding Source under the terms of this License,
		in one of these ways:
		
		    a) Convey the object code in, or embodied in, a physical product
		    (including a physical distribution medium), accompanied by the
		    Corresponding Source fixed on a durable physical medium
		    customarily used for software interchange.
		
		    b) Convey the object code in, or embodied in, a physical product
		    (including a physical distribution medium), accompanied by a
		    written offer, valid for at least three years and valid for as
		    long as you offer spare parts or customer support for that product
		    model, to give anyone who possesses the object code either (1) a
		    copy of the Corresponding Source for all the software in the
		    product that is covered by this License, on a durable physical
		    medium customarily used for software interchange, for a price no
		    more than your reasonable cost of physically performing this
		    conveying of source, or (2) access to copy the
		    Corresponding Source from a network server at no charge.
		
		    c) Convey individual copies of the object code with a copy of the
		    written offer to provide the Corresponding Source.  This
		    alternative is allowed only occasionally and noncommercially, and
		    only if you received the object code with such an offer, in accord
		    with subsection 6b.
		
		    d) Convey the object code by offering access from a designated
		    place (gratis or for a charge), and offer equivalent access to the
		    Corresponding Source in the same way through the same place at no
		    further charge.  You need not require recipients to copy the
		    Corresponding Source along with the object code.  If the place to
		    copy the object code is a network server, the Corresponding Source
		    may be on a different server (operated by you or a third party)
		    that supports equivalent copying facilities, provided you maintain
		    clear directions next to the object code saying where to find the
		    Corresponding Source.  Regardless of what server hosts the
		    Corresponding Source, you remain obligated to ensure that it is
		    available for as long as needed to satisfy these requirements.
		
		    e) Convey the object code using peer-to-peer transmission, provided
		    you inform other peers where the object code and Corresponding
		    Source of the work are being offered to the general public at no
		    charge under subsection 6d.
		
		  A separable portion of the object code, whose source code is excluded
		from the Corresponding Source as a System Library, need not be
		included in conveying the object code work.
		
		  A "User Product" is either (1) a "consumer product", which means any
		tangible personal property which is normally used for personal, family,
		or household purposes, or (2) anything designed or sold for incorporation
		into a dwelling.  In determining whether a product is a consumer product,
		doubtful cases shall be resolved in favor of coverage.  For a particular
		product received by a particular user, "normally used" refers to a
		typical or common use of that class of product, regardless of the status
		of the particular user or of the way in which the particular user
		actually uses, or expects or is expected to use, the product.  A product
		is a consumer product regardless of whether the product has substantial
		commercial, industrial or non-consumer uses, unless such uses represent
		the only significant mode of use of the product.
		
		  "Installation Information" for a User Product means any methods,
		procedures, authorization keys, or other information required to install
		and execute modified versions of a covered work in that User Product from
		a modified version of its Corresponding Source.  The information must
		suffice to ensure that the continued functioning of the modified object
		code is in no case prevented or interfered with solely because
		modification has been made.
		
		  If you convey an object code work under this section in, or with, or
		specifically for use in, a User Product, and the conveying occurs as
		part of a transaction in which the right of possession and use of the
		User Product is transferred to the recipient in perpetuity or for a
		fixed term (regardless of how the transaction is characterized), the
		Corresponding Source conveyed under this section must be accompanied
		by the Installation Information.  But this requirement does not apply
		if neither you nor any third party retains the ability to install
		modified object code on the User Product (for example, the work has
		been installed in ROM).
		
		  The requirement to provide Installation Information does not include a
		requirement to continue to provide support service, warranty, or updates
		for a work that has been modified or installed by the recipient, or for
		the User Product in which it has been modified or installed.  Access to a
		network may be denied when the modification itself materially and
		adversely affects the operation of the network or violates the rules and
		protocols for communication across the network.
		
		  Corresponding Source conveyed, and Installation Information provided,
		in accord with this section must be in a format that is publicly
		documented (and with an implementation available to the public in
		source code form), and must require no special password or key for
		unpacking, reading or copying.
		
		  7. Additional Terms.
		
		  "Additional permissions" are terms that supplement the terms of this
		License by making exceptions from one or more of its conditions.
		Additional permissions that are applicable to the entire Program shall
		be treated as though they were included in this License, to the extent
		that they are valid under applicable law.  If additional permissions
		apply only to part of the Program, that part may be used separately
		under those permissions, but the entire Program remains governed by
		this License without regard to the additional permissions.
		
		  When you convey a copy of a covered work, you may at your option
		remove any additional permissions from that copy, or from any part of
		it.  (Additional permissions may be written to require their own
		removal in certain cases when you modify the work.)  You may place
		additional permissions on material, added by you to a covered work,
		for which you have or can give appropriate copyright permission.
		
		  Notwithstanding any other provision of this License, for material you
		add to a covered work, you may (if authorized by the copyright holders of
		that material) supplement the terms of this License with terms:
		
		    a) Disclaiming warranty or limiting liability differently from the
		    terms of sections 15 and 16 of this License; or
		
		    b) Requiring preservation of specified reasonable legal notices or
		    author attributions in that material or in the Appropriate Legal
		    Notices displayed by works containing it; or
		
		    c) Prohibiting misrepresentation of the origin of that material, or
		    requiring that modified versions of such material be marked in
		    reasonable ways as different from the original version; or
		
		    d) Limiting the use for publicity purposes of names of licensors or
		    authors of the material; or
		
		    e) Declining to grant rights under trademark law for use of some
		    trade names, trademarks, or service marks; or
		
		    f) Requiring indemnification of licensors and authors of that
		    material by anyone who conveys the material (or modified versions of
		    it) with contractual assumptions of liability to the recipient, for
		    any liability that these contractual assumptions directly impose on
		    those licensors and authors.
		
		  All other non-permissive additional terms are considered "further
		restrictions" within the meaning of section 10.  If the Program as you
		received it, or any part of it, contains a notice stating that it is
		governed by this License along with a term that is a further
		restriction, you may remove that term.  If a license document contains
		a further restriction but permits relicensing or conveying under this
		License, you may add to a covered work material governed by the terms
		of that license document, provided that the further restriction does
		not survive such relicensing or conveying.
		
		  If you add terms to a covered work in accord with this section, you
		must place, in the relevant source files, a statement of the
		additional terms that apply to those files, or a notice indicating
		where to find the applicable terms.
		
		  Additional terms, permissive or non-permissive, may be stated in the
		form of a separately written license, or stated as exceptions;
		the above requirements apply either way.
		
		  8. Termination.
		
		  You may not propagate or modify a covered work except as expressly
		provided under this License.  Any attempt otherwise to propagate or
		modify it is void, and will automatically terminate your rights under
		this License (including any patent licenses granted under the third
		paragraph of section 11).
		
		  However, if you cease all violation of this License, then your
		license from a particular copyright holder is reinstated (a)
		provisionally, unless and until the copyright holder explicitly and
		finally terminates your license, and (b) permanently, if the copyright
		holder fails to notify you of the violation by some reasonable means
		prior to 60 days after the cessation.
		
		  Moreover, your license from a particular copyright holder is
		reinstated permanently if the copyright holder notifies you of the
		violation by some reasonable means, this is the first time you have
		received notice of violation of this License (for any work) from that
		copyright holder, and you cure the violation prior to 30 days after
		your receipt of the notice.
		
		  Termination of your rights under this section does not terminate the
		licenses of parties who have received copies or rights from you under
		this License.  If your rights have been terminated and not permanently
		reinstated, you do not qualify to receive new licenses for the same
		material under section 10.
		
		  9. Acceptance Not Required for Having Copies.
		
		  You are not required to accept this License in order to receive or
		run a copy of the Program.  Ancillary propagation of a covered work
		occurring solely as a consequence of using peer-to-peer transmission
		to receive a copy likewise does not require acceptance.  However,
		nothing other than this License grants you permission to propagate or
		modify any covered work.  These actions infringe copyright if you do
		not accept this License.  Therefore, by modifying or propagating a
		covered work, you indicate your acceptance of this License to do so.
		
		  10. Automatic Licensing of Downstream Recipients.
		
		  Each time you convey a covered work, the recipient automatically
		receives a license from the original licensors, to run, modify and
		propagate that work, subject to this License.  You are not responsible
		for enforcing compliance by third parties with this License.
		
		  An "entity transaction" is a transaction transferring control of an
		organization, or substantially all assets of one, or subdividing an
		organization, or merging organizations.  If propagation of a covered
		work results from an entity transaction, each party to that
		transaction who receives a copy of the work also receives whatever
		licenses to the work the party's predecessor in interest had or could
		give under the previous paragraph, plus a right to possession of the
		Corresponding Source of the work from the predecessor in interest, if
		the predecessor has it or can get it with reasonable efforts.
		
		  You may not impose any further restrictions on the exercise of the
		rights granted or affirmed under this License.  For example, you may
		not impose a license fee, royalty, or other charge for exercise of
		rights granted under this License, and you may not initiate litigation
		(including a cross-claim or counterclaim in a lawsuit) alleging that
		any patent claim is infringed by making, using, selling, offering for
		sale, or importing the Program or any portion of it.
		
		  11. Patents.
		
		  A "contributor" is a copyright holder who authorizes use under this
		License of the Program or a work on which the Program is based.  The
		work thus licensed is called the contributor's "contributor version".
		
		  A contributor's "essential patent claims" are all patent claims
		owned or controlled by the contributor, whether already acquired or
		hereafter acquired, that would be infringed by some manner, permitted
		by this License, of making, using, or selling its contributor version,
		but do not include claims that would be infringed only as a
		consequence of further modification of the contributor version.  For
		purposes of this definition, "control" includes the right to grant
		patent sublicenses in a manner consistent with the requirements of
		this License.
		
		  Each contributor grants you a non-exclusive, worldwide, royalty-free
		patent license under the contributor's essential patent claims, to
		make, use, sell, offer for sale, import and otherwise run, modify and
		propagate the contents of its contributor version.
		
		  In the following three paragraphs, a "patent license" is any express
		agreement or commitment, however denominated, not to enforce a patent
		(such as an express permission to practice a patent or covenant not to
		sue for patent infringement).  To "grant" such a patent license to a
		party means to make such an agreement or commitment not to enforce a
		patent against the party.
		
		  If you convey a covered work, knowingly relying on a patent license,
		and the Corresponding Source of the work is not available for anyone
		to copy, free of charge and under the terms of this License, through a
		publicly available network server or other readily accessible means,
		then you must either (1) cause the Corresponding Source to be so
		available, or (2) arrange to deprive yourself of the benefit of the
		patent license for this particular work, or (3) arrange, in a manner
		consistent with the requirements of this License, to extend the patent
		license to downstream recipients.  "Knowingly relying" means you have
		actual knowledge that, but for the patent license, your conveying the
		covered work in a country, or your recipient's use of the covered work
		in a country, would infringe one or more identifiable patents in that
		country that you have reason to believe are valid.
		
		  If, pursuant to or in connection with a single transaction or
		arrangement, you convey, or propagate by procuring conveyance of, a
		covered work, and grant a patent license to some of the parties
		receiving the covered work authorizing them to use, propagate, modify
		or convey a specific copy of the covered work, then the patent license
		you grant is automatically extended to all recipients of the covered
		work and works based on it.
		
		  A patent license is "discriminatory" if it does not include within
		the scope of its coverage, prohibits the exercise of, or is
		conditioned on the non-exercise of one or more of the rights that are
		specifically granted under this License.  You may not convey a covered
		work if you are a party to an arrangement with a third party that is
		in the business of distributing software, under which you make payment
		to the third party based on the extent of your activity of conveying
		the work, and under which the third party grants, to any of the
		parties who would receive the covered work from you, a discriminatory
		patent license (a) in connection with copies of the covered work
		conveyed by you (or copies made from those copies), or (b) primarily
		for and in connection with specific products or compilations that
		contain the covered work, unless you entered into that arrangement,
		or that patent license was granted, prior to 28 March 2007.
		
		  Nothing in this License shall be construed as excluding or limiting
		any implied license or other defenses to infringement that may
		otherwise be available to you under applicable patent law.
		
		  12. No Surrender of Others' Freedom.
		
		  If conditions are imposed on you (whether by court order, agreement or
		otherwise) that contradict the conditions of this License, they do not
		excuse you from the conditions of this License.  If you cannot convey a
		covered work so as to satisfy simultaneously your obligations under this
		License and any other pertinent obligations, then as a consequence you may
		not convey it at all.  For example, if you agree to terms that obligate you
		to collect a royalty for further conveying from those to whom you convey
		the Program, the only way you could satisfy both those terms and this
		License would be to refrain entirely from conveying the Program.
		
		  13. Remote Network Interaction; Use with the GNU General Public License.
		
		  Notwithstanding any other provision of this License, if you modify the
		Program, your modified version must prominently offer all users
		interacting with it remotely through a computer network (if your version
		supports such interaction) an opportunity to receive the Corresponding
		Source of your version by providing access to the Corresponding Source
		from a network server at no charge, through some standard or customary
		means of facilitating copying of software.  This Corresponding Source
		shall include the Corresponding Source for any work covered by version 3
		of the GNU General Public License that is incorporated pursuant to the
		following paragraph.
		
		  Notwithstanding any other provision of this License, you have
		permission to link or combine any covered work with a work licensed
		under version 3 of the GNU General Public License into a single
		combined work, and to convey the resulting work.  The terms of this
		License will continue to apply to the part which is the covered work,
		but the work with which it is combined will remain governed by version
		3 of the GNU General Public License.
		
		  14. Revised Versions of this License.
		
		  The Free Software Foundation may publish revised and/or new versions of
		the GNU Affero General Public License from time to time.  Such new versions
		will be similar in spirit to the present version, but may differ in detail to
		address new problems or concerns.
		
		  Each version is given a distinguishing version number.  If the
		Program specifies that a certain numbered version of the GNU Affero General
		Public License "or any later version" applies to it, you have the
		option of following the terms and conditions either of that numbered
		version or of any later version published by the Free Software
		Foundation.  If the Program does not specify a version number of the
		GNU Affero General Public License, you may choose any version ever published
		by the Free Software Foundation.
		
		  If the Program specifies that a proxy can decide which future
		versions of the GNU Affero General Public License can be used, that proxy's
		public statement of acceptance of a version permanently authorizes you
		to choose that version for the Program.
		
		  Later license versions may give you additional or different
		permissions.  However, no additional obligations are imposed on any
		author or copyright holder as a result of your choosing to follow a
		later version.
		
		  15. Disclaimer of Warranty.
		
		  THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY
		APPLICABLE LAW.  EXCEPT WHEN OTHERWISE STATED IN WRITING THE COPYRIGHT
		HOLDERS AND/OR OTHER PARTIES PROVIDE THE PROGRAM "AS IS" WITHOUT WARRANTY
		OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO,
		THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR
		PURPOSE.  THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM
		IS WITH YOU.  SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF
		ALL NECESSARY SERVICING, REPAIR OR CORRECTION.
		
		  16. Limitation of Liability.
		
		  IN NO EVENT UNLESS REQUIRED BY APPLICABLE LAW OR AGREED TO IN WRITING
		WILL ANY COPYRIGHT HOLDER, OR ANY OTHER PARTY WHO MODIFIES AND/OR CONVEYS
		THE PROGRAM AS PERMITTED ABOVE, BE LIABLE TO YOU FOR DAMAGES, INCLUDING ANY
		GENERAL, SPECIAL, INCIDENTAL OR CONSEQUENTIAL DAMAGES ARISING OUT OF THE
		USE OR INABILITY TO USE THE PROGRAM (INCLUDING BUT NOT LIMITED TO LOSS OF
		DATA OR DATA BEING RENDERED INACCURATE OR LOSSES SUSTAINED BY YOU OR THIRD
		PARTIES OR A FAILURE OF THE PROGRAM TO OPERATE WITH ANY OTHER PROGRAMS),
		EVEN IF SUCH HOLDER OR OTHER PARTY HAS BEEN ADVISED OF THE POSSIBILITY OF
		SUCH DAMAGES.
		
		  17. Interpretation of Sections 15 and 16.
		
		  If the disclaimer of warranty and limitation of liability provided
		above cannot be given local legal effect according to their terms,
		reviewing courts shall apply local law that most closely approximates
		an absolute waiver of all civil liability in connection with the
		Program, unless a warranty or assumption of liability accompanies a
		copy of the Program in return for a fee.
		
		                     END OF TERMS AND CONDITIONS
		
		            How to Apply These Terms to Your New Programs
		
		  If you develop a new program, and you want it to be of the greatest
		possible use to the public, the best way to achieve this is to make it
		free software which everyone can redistribute and change under these terms.
		
		  To do so, attach the following notices to the program.  It is safest
		to attach them to the start of each source file to most effectively
		state the exclusion of warranty; and each file should have at least
		the "copyright" line and a pointer to where the full notice is found.
		
		    <one line to give the program's name and a brief idea of what it does.>
		    Copyright (C) <year>  <name of author>
		
		    This program is free software: you can redistribute it and/or modify
		    it under the terms of the GNU Affero General Public License as published by
		    the Free Software Foundation, either version 3 of the License, or
		    (at your option) any later version.
		
		    This program is distributed in the hope that it will be useful,
		    but WITHOUT ANY WARRANTY; without even the implied warranty of
		    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
		    GNU Affero General Public License for more details.
		
		    You should have received a copy of the GNU Affero General Public License
		    along with this program.  If not, see <http://www.gnu.org/licenses/>.
		
		Also add information on how to contact you by electronic and paper mail.
		
		  If your software can interact with users remotely through a computer
		network, you should also make sure that it provides a way for users to
		get its source.  For example, if your program is a web application, its
		interface could display a "Source" link that leads users to an archive
		of the code.  There are many ways you could offer source, and different
		solutions will be better for different programs; see section 13 for the
		specific requirements.
		
		  You should also get your employer (if you work as a programmer) or school,
		if any, to sign a "copyright disclaimer" for the program, if necessary.
		For more information on this, and how to apply and follow the GNU AGPL, see
		<http://www.gnu.org/licenses/>.
		
		
	#tag EndNote

	#tag Note, Name = How to use
		Reference: https://www.virustotal.com/documentation/public-api/
		
		All public functions of this module correspond to an actions available 
		through the public VirusTotal API v.2. All interactions with the API require an
		API key. 
		
		The Virus Total API returns JSON. Luckily, REALstudio has shipped with built-in 
		JSON support since RS2011r2. All the public functions of this module, therefore, 
		return JSONItems (or Nil, on error.)
		
		If the returned JSONItem was not Nil, then LastResponseCode and 
		LastResponseVerbose correspond to the response_code and verbose_msg members 
		of Virus Total's response. 
		
		If the returned JSONItem was Nil, and it was because of a socket error, then 
		LastResponseCode and LastResponseVerbose correspond to the RB socket error 
		number and a brief error message. 
		
		If the returned JSONItem was Nil, and it was because of a JSON error, then 
		LastResponseCode is VTAPI.INVALID_RESPONSE and LastResponseVerbose is a 
		brief error message.
	#tag EndNote


	#tag Property, Flags = &h1
		Protected LastResponseCode As Integer
	#tag EndProperty

	#tag Property, Flags = &h1
		Protected LastResponseVerbose As String
	#tag EndProperty


	#tag Constant, Name = INVALID_RESPONSE, Type = Double, Dynamic = False, Default = \"255", Scope = Protected
	#tag EndConstant

	#tag Constant, Name = MIMEBoundary, Type = String, Dynamic = False, Default = \"--------0xKhTmLMIMEBoundary", Scope = Private
	#tag EndConstant

	#tag Constant, Name = VT_Get_File, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/file/report", Scope = Private
	#tag EndConstant

	#tag Constant, Name = VT_Get_URL, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/url/scan", Scope = Private
	#tag EndConstant

	#tag Constant, Name = VT_Put_Comment, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/comments/put", Scope = Private
	#tag EndConstant

	#tag Constant, Name = VT_Rescan_File, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/file/rescan", Scope = Private
	#tag EndConstant

	#tag Constant, Name = VT_Submit_File, Type = String, Dynamic = False, Default = \"www.virustotal.com/vtapi/v2/file/scan", Scope = Private
	#tag EndConstant


	#tag Enum, Name = ReportType, Type = Integer, Flags = &h1
		FileReport
		URLReport
	#tag EndEnum


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
