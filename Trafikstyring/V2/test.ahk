; test_vl := vognløbObj()

; test_data := Map(
;     "vognløbsnummer", "31200",
;     "vognløbsdato", A_Now,
;     "kørselsaftale", "3100",
;     "styresystem", "47"
; )

; test_vl.hent_data_vognløb_alt_obj(test_data)


test_data := [map("vognløbsnummer", "31200", "vognløbsdato", A_Now, "kørselsaftale", "3100", "styresystem", "47"), map("vognløbsnummer", "31201", "vognløbsdato", A_Now, "kørselsaftale", "3101", "styresystem", "47"), map("vognløbsnummer", "31202", "vognløbsdato", A_Now, "kørselsaftale", "3102", "styresystem", "47"), map("vognløbsnummer", "31203", "vognløbsdato", A_Now, "kørselsaftale", "3103", "styresystem", "47"), map("vognløbsnummer", "31204", "vognløbsdato", A_Now, "kørselsaftale", "3104", "styresystem", "47"), map("vognløbsnummer", "31205", "vognløbsdato", A_Now, "kørselsaftale", "3105", "styresystem", "47")]

vl_array := []

for index, data in test_data
{

    vl_array.Push(vognløbObj())
    vl_array[index].hent_data_vognløb_alt_obj(test_data[index])
    vl_array[index].gv_tjek()


}

for i,e in vl_array
    MsgBox e.vognløbsnummer