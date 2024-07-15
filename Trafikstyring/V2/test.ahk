test_vl := vognløbObj()

test_data := Map(
    "vognløbsnummer", "31200",
    "vognløbsdato", A_Now,
    "kørselsaftale", "3100",
    "styresystem", "47"
)

test_vl.hent_data_vognløb_alt_obj(test_data)