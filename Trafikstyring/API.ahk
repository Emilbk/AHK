#Requires AutoHotkey v1.1

; Version: 2023.07.05.2
; https://gist.github.com/e6062286ac7f4c35b612d3a53535cc2a
; Usage and examples: https://redd.it/mcjj4s
; Testing: http://httpbin.org/ | http://httpbun.org/ | http://ptsv2.com/

WinHttpRequest(oOptions := "") {
    return new WinHttpRequest(oOptions)
}

class WinHttpRequest extends WinHttpRequest.Functor {

    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")

    ;#region: Meta

    __New(oOptions := "") {
        static HTTPREQUEST_PROXYSETTING_DEFAULT := 0, HTTPREQUEST_PROXYSETTING_DIRECT := 1, HTTPREQUEST_PROXYSETTING_PROXY := 2, EnableCertificateRevocationCheck := 18, SslErrorIgnoreFlags := 4, SslErrorFlag_Ignore_All := 13056, SecureProtocols := 9, WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_3 := 8192, WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2 := 2048, UserAgentString := 0
        if (!IsObject(oOptions)) {
            oOptions := {}
        }
        if (!oOptions.HasKey("Proxy") || !oOptions.Proxy) {
            this.whr.SetProxy(HTTPREQUEST_PROXYSETTING_DEFAULT)
        } else if (oOptions.Proxy = "DIRECT") {
            this.whr.SetProxy(HTTPREQUEST_PROXYSETTING_DIRECT)
        } else {
            this.whr.SetProxy(HTTPREQUEST_PROXYSETTING_PROXY, oOptions.Proxy)
        }
        if (oOptions.HasKey("Revocation")) {
            this.whr.Option[EnableCertificateRevocationCheck] := !!oOptions.Revocation
        } else {
            this.whr.Option[EnableCertificateRevocationCheck] := true
        }
        if (oOptions.HasKey("SslError")) {
            if (oOptions.SslError = false) {
                this.whr.Option[SslErrorIgnoreFlags] := SslErrorFlag_Ignore_All
            }
        }
        if (!oOptions.HasKey("TLS")) {
            this.whr.Option[SecureProtocols] := WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_3 | WINHTTP_FLAG_SECURE_PROTOCOL_TLS1_2
        } else {
            this.whr.Option[SecureProtocols] := oOptions.TLS
        }
        if (oOptions.HasKey("UA")) {
            this.whr.Option[UserAgentString] := oOptions.UA
        }
    }
    ;#endregion

    ;#region: Static

    EncodeUri(sUri) {
        return this._EncodeDecode(sUri, true, false)
    }

    EncodeUriComponent(sComponent) {
        return this._EncodeDecode(sComponent, true, true)
    }

    DecodeUri(sUri) {
        return this._EncodeDecode(sUri, false, false)
    }

    DecodeUriComponent(sComponent) {
        return this._EncodeDecode(sComponent, false, true)
    }

    ObjToQuery(oData) {
        if (!IsObject(oData)) {
            return oData
        }
        out := ""
        for key, val in oData {
            out .= this._EncodeDecode(key, true, true) "="
            out .= this._EncodeDecode(val, true, true) "&"
        }
        return RTrim(out, "&")
    }

    QueryToObj(sData) {
        if (IsObject(sData)) {
            return sData
        }
        sData := LTrim(sData, "?")
        obj := {}
        for _, part in StrSplit(sData, "&") {
            pair := StrSplit(part, "=", "", 2)
            key := this._EncodeDecode(pair[1], false, true)
            val := this._EncodeDecode(pair[2], false, true)
            obj[key] := val
        }
        return obj
    }
    ;#endregion

    ;#region: Public

    Request(sMethod, sUrl, mBody := "", oHeaders := false, oOptions := false) {
        if (this.whr = "") {
            throw Exception("Not initialized.", -1)
        }
        sMethod := Format("{:U}", sMethod) ; CONNECT not supported
        if !(sMethod ~= "^(DELETE|GET|HEAD|OPTIONS|PATCH|POST|PUT|TRACE)$") {
            throw Exception("Invalid HTTP verb.", -1, sMethod)
        }
        if !(sUrl := Trim(sUrl)) {
            throw Exception("Empty URL.", -1)
        }
        if (!IsObject(oHeaders)) {
            oHeaders := {}
        }
        if (!IsObject(oOptions)) {
            oOptions := {}
        }
        if (sMethod = "POST") {
            multi := oOptions.HasKey("Multipart") ? !!oOptions.Multipart : false
            this._Post(mBody, oHeaders, multi)
        } else if (sMethod = "GET" && mBody) {
            sUrl := RTrim(sUrl, "&")
            sUrl .= InStr(sUrl, "?") ? "&" : "?"
            sUrl .= WinHttpRequest.ObjToQuery(mBody)
            mBody := ""
        }
        this.whr.Open(sMethod, sUrl, true)
        for key, val in oHeaders {
            this.whr.SetRequestHeader(key, val)
        }
        this.whr.Send(mBody)
        this.whr.WaitForResponse()
        if (oOptions.HasKey("Save")) {
            target := RegExReplace(oOptions.Save, "^\h*\*\h*", "", forceSave)
            if (this.whr.Status = 200 || forceSave) {
                this._Save(target)
            }
            return this.whr.Status
        }
        out := new WinHttpRequest._Response()
        out.Headers := this._Headers()
        out.Status := this.whr.Status
        out.Text := this._Text(oOptions.HasKey("Encoding") ? oOptions.Encoding : "")
        return out
    }
    ;#endregion

    ;#region: Private

    static _doc := ""

    _EncodeDecode(Text, bEncode, bComponent) {
        if (this._doc = "") {
            this._doc := ComObjCreate("HTMLFile")
            this._doc.write("<meta http-equiv='X-UA-Compatible' content='IE=Edge'>")
        }
        action := (bEncode ? "en" : "de") "codeURI" (bComponent ? "Component" : "")
        return ObjBindMethod(this._doc.parentWindow, action).Call(Text)
    }

    _Headers() {
        headers := this.whr.GetAllResponseHeaders()
        headers := RTrim(headers, "`r`n")
        out := {}
        for _, line in StrSplit(headers, "`n", "`r") {
            pair := StrSplit(line, ":", " ", 2)
            out[pair[1]] := pair[2]
        }
        return out
    }

    _Mime(Extension) {
        if (WinHttpRequest.MIME.HasKey(Extension)) {
            return WinHttpRequest.MIME[Extension]
        }
        return "application/octet-stream"
    }

    _MultiPart(ByRef Body) {
        static LMEM_ZEROINIT := 64, EOL := "`r`n"
        this._memLen := 0
        this._memPtr := DllCall("LocalAlloc", "UInt", LMEM_ZEROINIT, "UInt", 1)
        boundary := "----------WinHttpRequest-" A_NowUTC A_MSec
        for field, value in Body {
            this._MultiPartAdd(boundary, EOL, field, value)
        }
        this._MultipartStr("--" boundary "--" EOL)
        Body := ComObjArray(0x11, this._memLen)
        pvData := NumGet(ComObjValue(Body) + 8 + A_PtrSize, "Ptr")
        DllCall("RtlMoveMemory", "Ptr", pvData, "Ptr", this._memPtr, "UInt", this._memLen)
        DllCall("LocalFree", "Ptr", this._memPtr)
        return boundary
    }

    _MultiPartAdd(Boundary, EOL, Field, Value) {
        if (!IsObject(Value)) {
            str := "--" Boundary
            str .= EOL
            str .= "Content-Disposition: form-data; name=""" Field """"
            str .= EOL
            str .= EOL
            str .= Value
            str .= EOL
            this._MultipartStr(str)
            return
        }
        for _, path in Value {
            SplitPath path, filename, , ext
            str := "--" Boundary
            str .= EOL
            str .= "Content-Disposition: form-data; name=""" Field """; filename=""" filename """"
            str .= EOL
            str .= "Content-Type: " this._Mime(ext)
            str .= EOL
            str .= EOL
            this._MultipartStr(str)
            this._MultipartFile(path)
            this._MultipartStr(EOL)
        }
    }

    _MultipartFile(Path) {
        static LHND := 66
        try {
            oFile := FileOpen(Path, 0x0)
        } catch {
            throw Exception("Couldn't open file for reading.", -1, Path)
        }
        this._memLen += oFile.Length
        this._memPtr := DllCall("LocalReAlloc", "Ptr", this._memPtr, "UInt", this._memLen, "UInt", LHND)
        oFile.RawRead(this._memPtr + this._memLen - oFile.length, oFile.length)
    }

    _MultipartStr(Text) {
        static LHND := 66
        size := StrPut(Text, "UTF-8") - 1
        this._memLen += size
        this._memPtr := DllCall("LocalReAlloc", "Ptr", this._memPtr, "UInt", this._memLen, "UInt", LHND)
        StrPut(Text, this._memPtr + this._memLen - size, size, "UTF-8")
    }

    _Post(ByRef Body, ByRef Headers, bMultipart) {
        isMultipart := 0
        for _, value in Body {
            isMultipart += !!IsObject(value)
        }
        if (isMultipart || bMultipart) {
            Body := WinHttpRequest.QueryToObj(Body)
            boundary := this._MultiPart(Body)
            Headers["Content-Type"] := "multipart/form-data; boundary=""" boundary """"
        } else {
            Body := WinHttpRequest.ObjToQuery(Body)
            if (!Headers.HasKey("Content-Type")) {
                Headers["Content-Type"] := "application/x-www-form-urlencoded"
            }
        }
    }

    _Save(Path) {
        arr := this.whr.ResponseBody
        pData := NumGet(ComObjValue(arr) + 8 + A_PtrSize, "Ptr")
        length := arr.MaxIndex() + 1
        FileOpen(Path, 0x1).RawWrite(pData + 0, length)
    }

    _Text(Encoding) {
        response := ""
        try response := this.whr.ResponseText
        if (response = "" || Encoding != "") {
            try {
                arr := this.whr.ResponseBody
                pData := NumGet(ComObjValue(arr) + 8 + A_PtrSize, "Ptr")
                length := arr.MaxIndex() + 1
                response := StrGet(pData, length, Encoding)
            }
        }
        return response
    }

    class Functor {

        __Call(Method, Parameters*) {
            return this.Request(Method, Parameters*)
        }

    }

    class _Response {

        Json {
            get {
                method := Json.HasKey("parse") ? "parse" : "Load"
                oJson := ObjBindMethod(Json, method, this.Text).Call()
                ObjRawSet(this, "Json", oJson)
                return oJson
            }
        }

    }

    ;#endregion

    class MIME {
        static 7z := "application/x-7z-compressed"
        static gif := "image/gif"
        static jpg := "image/jpeg"
        static json := "application/json"
        static png := "image/png"
        static zip := "application/zip"
    }

}


/**
 * Lib: JSON.ahk
 *     JSON lib for AutoHotkey.
 * Version:
 *     v2.1.3 [updated 04/18/2016 (MM/DD/YYYY)]
 * License:
 *     WTFPL [http://wtfpl.net/]
 * Requirements:
 *     Latest version of AutoHotkey (v1.1+ or v2.0-a+)
 * Installation:
 *     Use #Include JSON.ahk or copy into a function library folder and then
 *     use #Include <JSON>
 * Links:
 *     GitHub:     - https://github.com/cocobelgica/AutoHotkey-JSON
 *     Forum Topic - http://goo.gl/r0zI8t
 *     Email:      - cocobelgica <at> gmail <dot> com
 */


/**
 * Class: JSON
 *     The JSON object contains methods for parsing JSON and converting values
 *     to JSON. Callable - NO; Instantiable - YES; Subclassable - YES;
 *     Nestable(via #Include) - NO.
 * Methods:
 *     Load() - see relevant documentation before method definition header
 *     Dump() - see relevant documentation before method definition header
 */
 class JSON
 {
     /**
      * Method: Load
      *     Parses a JSON string into an AHK value
      * Syntax:
      *     value := JSON.Load( text [, reviver ] )
      * Parameter(s):
      *     value      [retval] - parsed value
      *     text    [in, ByRef] - JSON formatted string
      *     reviver   [in, opt] - function object, similar to JavaScript's
      *                           JSON.parse() 'reviver' parameter
      */
     class Load extends JSON.Functor
     {
         Call(self, ByRef text, reviver:="")
         {
             this.rev := IsObject(reviver) ? reviver : false
         ; Object keys(and array indices) are temporarily stored in arrays so that
         ; we can enumerate them in the order they appear in the document/text instead
         ; of alphabetically. Skip if no reviver function is specified.
             this.keys := this.rev ? {} : false
 
             static quot := Chr(34), bashq := "\" . quot
                  , json_value := quot . "{[01234567890-tfn"
                  , json_value_or_array_closing := quot . "{[]01234567890-tfn"
                  , object_key_or_object_closing := quot . "}"
 
             key := ""
             is_key := false
             root := {}
             stack := [root]
             next := json_value
             pos := 0
 
             while ((ch := SubStr(text, ++pos, 1)) != "") {
                 if InStr(" `t`r`n", ch)
                     continue
                 if !InStr(next, ch, 1)
                     this.ParseError(next, text, pos)
 
                 holder := stack[1]
                 is_array := holder.IsArray
 
                 if InStr(",:", ch) {
                     next := (is_key := !is_array && ch == ",") ? quot : json_value
 
                 } else if InStr("}]", ch) {
                     ObjRemoveAt(stack, 1)
                     next := stack[1]==root ? "" : stack[1].IsArray ? ",]" : ",}"
 
                 } else {
                     if InStr("{[", ch) {
                     ; Check if Array() is overridden and if its return value has
                     ; the 'IsArray' property. If so, Array() will be called normally,
                     ; otherwise, use a custom base object for arrays
                         static json_array := Func("Array").IsBuiltIn || ![].IsArray ? {IsArray: true} : 0
                     
                     ; sacrifice readability for minor(actually negligible) performance gain
                         (ch == "{")
                             ? ( is_key := true
                               , value := {}
                               , next := object_key_or_object_closing )
                         ; ch == "["
                             : ( value := json_array ? new json_array : []
                               , next := json_value_or_array_closing )
                         
                         ObjInsertAt(stack, 1, value)
 
                         if (this.keys)
                             this.keys[value] := []
                     
                     } else {
                         if (ch == quot) {
                             i := pos
                             while (i := InStr(text, quot,, i+1)) {
                                 value := StrReplace(SubStr(text, pos+1, i-pos-1), "\\", "\u005c")
 
                                 static tail := A_AhkVersion<"2" ? 0 : -1
                                 if (SubStr(value, tail) != "\")
                                     break
                             }
 
                             if (!i)
                                 this.ParseError("'", text, pos)
 
                               value := StrReplace(value,  "\/",  "/")
                             , value := StrReplace(value, bashq, quot)
                             , value := StrReplace(value,  "\b", "`b")
                             , value := StrReplace(value,  "\f", "`f")
                             , value := StrReplace(value,  "\n", "`n")
                             , value := StrReplace(value,  "\r", "`r")
                             , value := StrReplace(value,  "\t", "`t")
 
                             pos := i ; update pos
                             
                             i := 0
                             while (i := InStr(value, "\",, i+1)) {
                                 if !(SubStr(value, i+1, 1) == "u")
                                     this.ParseError("\", text, pos - StrLen(SubStr(value, i+1)))
 
                                 uffff := Abs("0x" . SubStr(value, i+2, 4))
                                 if (A_IsUnicode || uffff < 0x100)
                                     value := SubStr(value, 1, i-1) . Chr(uffff) . SubStr(value, i+6)
                             }
 
                             if (is_key) {
                                 key := value, next := ":"
                                 continue
                             }
                         
                         } else {
                             value := SubStr(text, pos, i := RegExMatch(text, "[\]\},\s]|$",, pos)-pos)
 
                             static number := "number", integer :="integer"
                             if value is %number%
                             {
                                 if value is %integer%
                                     value += 0
                             }
                             else if (value == "true" || value == "false")
                                 value := %value% + 0
                             else if (value == "null")
                                 value := ""
                             else
                             ; we can do more here to pinpoint the actual culprit
                             ; but that's just too much extra work.
                                 this.ParseError(next, text, pos, i)
 
                             pos += i-1
                         }
 
                         next := holder==root ? "" : is_array ? ",]" : ",}"
                     } ; If InStr("{[", ch) { ... } else
 
                     is_array? key := ObjPush(holder, value) : holder[key] := value
 
                     if (this.keys && this.keys.HasKey(holder))
                         this.keys[holder].Push(key)
                 }
             
             } ; while ( ... )
 
             return this.rev ? this.Walk(root, "") : root[""]
         }
 
         ParseError(expect, ByRef text, pos, len:=1)
         {
             static quot := Chr(34), qurly := quot . "}"
             
             line := StrSplit(SubStr(text, 1, pos), "`n", "`r").Length()
             col := pos - InStr(text, "`n",, -(StrLen(text)-pos+1))
             msg := Format("{1}`n`nLine:`t{2}`nCol:`t{3}`nChar:`t{4}"
             ,     (expect == "")     ? "Extra data"
                 : (expect == "'")    ? "Unterminated string starting at"
                 : (expect == "\")    ? "Invalid \escape"
                 : (expect == ":")    ? "Expecting ':' delimiter"
                 : (expect == quot)   ? "Expecting object key enclosed in double quotes"
                 : (expect == qurly)  ? "Expecting object key enclosed in double quotes or object closing '}'"
                 : (expect == ",}")   ? "Expecting ',' delimiter or object closing '}'"
                 : (expect == ",]")   ? "Expecting ',' delimiter or array closing ']'"
                 : InStr(expect, "]") ? "Expecting JSON value or array closing ']'"
                 :                      "Expecting JSON value(string, number, true, false, null, object or array)"
             , line, col, pos)
 
             static offset := A_AhkVersion<"2" ? -3 : -4
             throw Exception(msg, offset, SubStr(text, pos, len))
         }
 
         Walk(holder, key)
         {
             value := holder[key]
             if IsObject(value) {
                 for i, k in this.keys[value] {
                     ; check if ObjHasKey(value, k) ??
                     v := this.Walk(value, k)
                     if (v != JSON.Undefined)
                         value[k] := v
                     else
                         ObjDelete(value, k)
                 }
             }
             
             return this.rev.Call(holder, key, value)
         }
     }
 
     /**
      * Method: Dump
      *     Converts an AHK value into a JSON string
      * Syntax:
      *     str := JSON.Dump( value [, replacer, space ] )
      * Parameter(s):
      *     str        [retval] - JSON representation of an AHK value
      *     value          [in] - any value(object, string, number)
      *     replacer  [in, opt] - function object, similar to JavaScript's
      *                           JSON.stringify() 'replacer' parameter
      *     space     [in, opt] - similar to JavaScript's JSON.stringify()
      *                           'space' parameter
      */
     class Dump extends JSON.Functor
     {
         Call(self, value, replacer:="", space:="")
         {
             this.rep := IsObject(replacer) ? replacer : ""
 
             this.gap := ""
             if (space) {
                 static integer := "integer"
                 if space is %integer%
                     Loop, % ((n := Abs(space))>10 ? 10 : n)
                         this.gap .= " "
                 else
                     this.gap := SubStr(space, 1, 10)
 
                 this.indent := "`n"
             }
 
             return this.Str({"": value}, "")
         }
 
         Str(holder, key)
         {
             value := holder[key]
 
             if (this.rep)
                 value := this.rep.Call(holder, key, ObjHasKey(holder, key) ? value : JSON.Undefined)
 
             if IsObject(value) {
             ; Check object type, skip serialization for other object types such as
             ; ComObject, Func, BoundFunc, FileObject, RegExMatchObject, Property, etc.
                 static type := A_AhkVersion<"2" ? "" : Func("Type")
                 if (type ? type.Call(value) == "Object" : ObjGetCapacity(value) != "") {
                     if (this.gap) {
                         stepback := this.indent
                         this.indent .= this.gap
                     }
 
                     is_array := value.IsArray
                 ; Array() is not overridden, rollback to old method of
                 ; identifying array-like objects. Due to the use of a for-loop
                 ; sparse arrays such as '[1,,3]' are detected as objects({}). 
                     if (!is_array) {
                         for i in value
                             is_array := i == A_Index
                         until !is_array
                     }
 
                     str := ""
                     if (is_array) {
                         Loop, % value.Length() {
                             if (this.gap)
                                 str .= this.indent
                             
                             v := this.Str(value, A_Index)
                             str .= (v != "") ? v . "," : "null,"
                         }
                     } else {
                         colon := this.gap ? ": " : ":"
                         for k in value {
                             v := this.Str(value, k)
                             if (v != "") {
                                 if (this.gap)
                                     str .= this.indent
 
                                 str .= this.Quote(k) . colon . v . ","
                             }
                         }
                     }
 
                     if (str != "") {
                         str := RTrim(str, ",")
                         if (this.gap)
                             str .= stepback
                     }
 
                     if (this.gap)
                         this.indent := stepback
 
                     return is_array ? "[" . str . "]" : "{" . str . "}"
                 }
             
             } else ; is_number ? value : "value"
                 return ObjGetCapacity([value], 1)=="" ? value : this.Quote(value)
         }
 
         Quote(string)
         {
             static quot := Chr(34), bashq := "\" . quot
 
             if (string != "") {
                   string := StrReplace(string,  "\",  "\\")
                 ; , string := StrReplace(string,  "/",  "\/") ; optional in ECMAScript
                 , string := StrReplace(string, quot, bashq)
                 , string := StrReplace(string, "`b",  "\b")
                 , string := StrReplace(string, "`f",  "\f")
                 , string := StrReplace(string, "`n",  "\n")
                 , string := StrReplace(string, "`r",  "\r")
                 , string := StrReplace(string, "`t",  "\t")
 
                 static rx_escapable := A_AhkVersion<"2" ? "O)[^\x20-\x7e]" : "[^\x20-\x7e]"
                 while RegExMatch(string, rx_escapable, m)
                     string := StrReplace(string, m.Value, Format("\u{1:04x}", Ord(m.Value)))
             }
 
             return quot . string . quot
         }
     }
 
     /**
      * Property: Undefined
      *     Proxy for 'undefined' type
      * Syntax:
      *     undefined := JSON.Undefined
      * Remarks:
      *     For use with reviver and replacer functions since AutoHotkey does not
      *     have an 'undefined' type. Returning blank("") or 0 won't work since these
      *     can't be distnguished from actual JSON values. This leaves us with objects.
      *     Replacer() - the caller may return a non-serializable AHK objects such as
      *     ComObject, Func, BoundFunc, FileObject, RegExMatchObject, and Property to
      *     mimic the behavior of returning 'undefined' in JavaScript but for the sake
      *     of code readability and convenience, it's better to do 'return JSON.Undefined'.
      *     Internally, the property returns a ComObject with the variant type of VT_EMPTY.
      */
     Undefined[]
     {
         get {
             static empty := {}, vt_empty := ComObject(0, &empty, 1)
             return vt_empty
         }
     }
 
     class Functor
     {
         __Call(method, ByRef arg, args*)
         {
         ; When casting to Call(), use a new instance of the "function object"
         ; so as to avoid directly storing the properties(used across sub-methods)
         ; into the "function object" itself.
             if IsObject(method)
                 return (new this).Call(method, arg, args*)
             else if (method == "")
                 return (new this).Call(arg, args*)
         }
     }
 }
;  "Authorization": "5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f",
json_str =
 (
 {
    "metrics": "distance",
    "sources": "0",
    "units": "km",
     "locations": [
        "9.70093,48.477473",
        "9.207916,49.153868",
        "37.573242,55.801281",
        "115.663757,38.106467"
    ]
 }
 )
 
 parsed := JSON.Load(json_str)



http := WinHttpRequest(oOptions)

gade := "møllevangen 23"
postnr := "8310"
komm := "Århus"

InputBox, gade, Gadenavn og nr
InputBox, postnr, Postnr
InputBox, kommune, Kommune 

endpoint := "https://api.openrouteservice.org/geocode/search?api_key=5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f&text=" gade " " postnr " " komm "&boundary.country=DK"
response := http.GET(endpoint)

data := JSON.Load(response.text)
geo2 := data.bbox[3]
geo1 := data.bbox[4]
clipboard := geo1 " " geo2
OutputDebug, % geo1
OutputDebug, % geo2


endpoint := "https://api.openrouteservice.org/v2/matrix/foot-walking"

; json_str = "{""locations"":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],""metrics"":[""distance""],""sources"":[0],""units":""km""}"

str := []
str.locations := [[geo2,geo1],[10.157957810640028,56.110729831443734],[37.573242,55.801281],[115.663757,38.106467]]
str.metrics := ["distance"]
str.sources := [0]
str.units := "km"

json_str := JSON.Dump(str)

; {"locations":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],"metrics":["distance"],"sources":[0],"units":"km"}


headers := []
headers["Content-Type"] := "application/json; charset=utf-8"
headers["Accept"] := "application/json, application/geo+json, application/gpx+xml, img/png"
headers["Authorization"] := "5b3ce3597851110001cf6248d9d7ce5fd9c74a9e8993312a027a2f4f"

; body := Map("{"locations":[[9.70093,48.477473],[9.207916,49.153868],[37.573242,55.801281],[115.663757,38.106467]],"metrics":["distance"],"sources":[0],"units":"km"}")
response_matrix := http.POST(endpoint, json_str, headers)

parsed_matrix := JSON.Load(response_matrix.text)

test := parsed_matrix.distances[1]

knudepunkt := []
knudepunkt.navn := ["Knudepunkt 1", "Knudepunkt 2", "Knudepunkt 3", "Knudepunkt 4"]
knudepunkt.geo := [geo2 "," geo1, "10.157957810640028,56.110729831443734", "37.573242,55.801281", "115.663757,38.106467"]


MsgBox, , Afstand, % "Nærmeste på " test[1] "km `n" test[2] " km `n" test[3] "`n" test[4]


; MsgBox,response.Text, "GET", 0x40040

; test := []
; test := [{"geocoding":{"version":"0.2","attribution":"https://openrouteservice.org/terms-of-service/#attribution-geocode","query":{"text":"møllevangen 23 8310 Århus","size":10,"private":false,"boundary.country":["DNK"],"lang":{"name":"English","iso6391":"en","iso6393":"eng","via":"default","defaulted":true},"querySize":20,"parser":"libpostal","parsed_text":{"street":"møllevangen","housenumber":"23","postalcode":"8310","city":"århus"}},"engine":{"name":"Pelias","author":"Mapzen","version":"1.0"},"timestamp":1696963717687},"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"Point","coordinates":[10.151441,56.102406]},"properties":{"id":"node/5665485241","gid":"openstreetmap:address:node/5665485241","layer":"address","source":"openstreetmap","source_id":"node/5665485241","name":"Møllevangen 23","housenumber":"23","street":"Møllevangen","postalcode":"8310","confidence":1,"match_type":"exact","accuracy":"point","country":"Denmark","country_gid":"whosonfirst:country:85633121","country_a":"DNK","region":"Central Jutland","region_gid":"whosonfirst:region:85682597","region_a":"MJ","localadmin":"Aarhus","localadmin_gid":"whosonfirst:localadmin:1394013977","locality":"Aarhus","locality_gid":"whosonfirst:locality:101749163","continent":"Europe","continent_gid":"whosonfirst:continent:102191581","label":"Møllevangen 23, Aarhus, MJ, Denmark"}}],"bbox":[10.151441,56.102406,10.151441,56.102406]}]
; MsgBox, 0x40040, "Get",% response.text,
MsgBox, 0x40040, "Get",% response_matrix.text,
; OutputDebug, % response.text
; geo := SubStr(response.Text, -2, -11)
; MsgBox, , , % geo,

; {"geocoding":{"version":"0.2","attribution":"https://openrouteservice.org/terms-of-service/#attribution-geocode","query":{"text":"møllevangen 23 8310 Århus","size":10,"private":false,"boundary.country":["DNK"],"lang":{"name":"English","iso6391":"en","iso6393":"eng","via":"default","defaulted":true},"querySize":20,"parser":"libpostal","parsed_text":{"street":"møllevangen","housenumber":"23","postalcode":"8310","city":"århus"}},"engine":{"name":"Pelias","author":"Mapzen","version":"1.0"},"timestamp":1696964025721},"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"Point","coordinates":[10.151441,56.102406]},"properties":{"id":"node/5665485241","gid":"openstreetmap:address:node/5665485241","layer":"address","source":"openstreetmap","source_id":"node/5665485241","name":"Møllevangen 23","housenumber":"23","street":"Møllevangen","postalcode":"8310","confidence":1,"match_type":"exact","accuracy":"point","country":"Denmark","country_gid":"whosonfirst:country:85633121","country_a":"DNK","region":"Central Jutland","region_gid":"whosonfirst:region:85682597","region_a":"MJ","localadmin":"Aarhus","localadmin_gid":"whosonfirst:localadmin:1394013977","locality":"Aarhus","locality_gid":"whosonfirst:locality:101749163","continent":"Europe","continent_gid":"whosonfirst:continent:102191581","label":"Møllevangen 23, Aarhus, MJ, Denmark"}}],"bbox":[10.151441,56.102406,10.151441,56.102406]}