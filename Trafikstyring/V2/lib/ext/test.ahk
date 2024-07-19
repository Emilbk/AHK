#Include "sql/SQLite.ahk"

db := SQLite("test.sql")

db.Exec('BEGIN TRANSACTION;')
db.Exec('CREATE TABLE IF NOT EXISTS test (id INTEGER PRIMARY KEY, name TEXT, value REAL)')
loop 20
	db.Exec('INSERT INTO test VALUES(' A_Index ', "name' A_Index '", "value' A_Index '");')
db.Exec('COMMIT TRANSACTION;')
table := db.Exec('SELECT * FROM test')
msgbox table.count

for row in table.rows
    {
        for header, value in row
            data .= header ' = ' value ', ' ; rows can be looped over as well
    
        data .= '`n'
    }
    
    msgbox data

return