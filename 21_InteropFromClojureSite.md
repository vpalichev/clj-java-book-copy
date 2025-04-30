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












