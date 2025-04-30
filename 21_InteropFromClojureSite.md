# Class access

- Classname
- Classname$InnerClass
- Classname/N
- primitive/N

- If a class is `import`ed in the namespace, it may be used without qualification. 

```clojure
String  ;; Classname
;;returns:  java.lang.String

;; All classes in java.lang are automatically imported to every namespace

(defn date? [d] (instance? java.util.Date d))  
;;returns:  #'user/date?

;; Fully-qualified class names are always valid

(.getEnclosingClass java.util.Map$Entry)  ;; Classname$InnerClass
;;returns:  java.util.Map

(.getComponentType String/1)
;;returns:  java.lang.String

;; Name with a single digit between 1 and 9, designates an array class of that component type and dimension. 
```

# Member access

- (.instanceMember instance args*) -> (.toUpperCase "fred") -> "FRED"/
- (.instanceMember Classname args*)/
- (.-instanceField instance)/
- (Classname/staticMethod args*)
- (Classname/.instanceMethod instance args*)
- Classname/staticField -> Math/PI -> 3.141592653589793/


```clojure
(.toUpperCase "fred")  ;; (.instanceMember instance args*)
;;returns:  "FRED"

(.getName String)  ;; (.instanceMember Classname args*)
;;returns:  "java.lang.String"

(.-x (java.awt.Point. 1 2))  ;; (.-instanceField instance)
;;returns:  1

(System/getProperty "java.vm.version")  ;; (Classname/staticMethod args*)
;;returns:  "1.6.0_07-b06-57"

Math/PI  ;; Classname/staticField
;;returns:  3.141592653589793
```







