class MyClass {
    ; Instance variable
    myVar := ""

    ; Constructor to initialize the instance variable
    __New(value) {
        this.myVar := value
    }

    ; Static method that takes an instance as a parameter
    static MyStaticMethod(instance) {
        MsgBox instance.myVar  ; Access the instance variable through the instance
    }
}

; Create an instance of the class
obj := MyClass("Hello, World!")

; Call the static method and pass the instance
MyClass.MyStaticMethod(obj)