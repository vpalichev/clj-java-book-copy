# Class access

- Classname
- Classname$InnerClass
- Classname/N
- primitive/N




- Symbols representing class names are resolved to the Class instance. 
- Inner or nested classes are separated from their outer class with a $. 
- Fully-qualified class names are always valid. 
- If a class is `import`ed in the namespace, it may be used without qualification. 
- All classes in java.lang are automatically imported to every namespace.



```clojure
String
-> java.lang.String

(defn date? [d] (instance? java.util.Date d))
-> #'user/date?

(.getEnclosingClass java.util.Map$Entry)
-> java.util.Map

(.getComponentType String/1)
-> java.lang.String
```


# Member access



- (.instanceMember instance args*) -> (.toUpperCase "fred") -> "FRED"
- (.instanceMember Classname args*)
- (.-instanceField instance)
- (Classname/staticMethod args*)
- (Classname/.instanceMethod instance args*)
- Classname/staticField -> Math/PI -> 3.141592653589793


```clojure
(.toUpperCase "fred")
-> "FRED"
(.getName String)
-> "java.lang.String"
(.-x (java.awt.Point. 1 2))
-> 1
(System/getProperty "java.vm.version")
-> "1.6.0_07-b06-57"
Math/PI

```







