#Requires AutoHotkey v2.0

; Create sample data array
data := [
    ["John", 25, "Engineer"],
    ["Alice", 30, "Doctor"],
    ["Bob", 22, "Artist"]
]

; Create the GUI
myGui := Gui()

; Create ListView with columns: Name, Age, and Occupation
myListView := myGui.Add("ListView", "w400 r10", ["Name", "Age", "Occupation"])

; Populate the ListView with data from the array
for i, row in data {
    
    myListView.Insert(A_Index, , row*) ; Add each row of data (using row* to unpack the array into parameters)
}

; Show the GUI
myGui.Show()
