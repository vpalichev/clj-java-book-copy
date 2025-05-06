# Samples from Jacob project


### Call hierarchy


**IDispatch::Invoke** method (oaidl.h):

https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke

```cpp
HRESULT Invoke(
  [in]      DISPID     dispIdMember, //Identifies the member. Use GetIDsOfNames or the object's documentation to obtain the dispatch identifier.
  [in]      REFIID     riid, //Reserved for future use. Must be IID_NULL.
  [in]      LCID       lcid, //The locale context in which to interpret arguments
  [in]      WORD       wFlags, //Flags describing the context of the Invoke call
  [in, out] DISPPARAMS *pDispParams, //Pointer to a DISPPARAMS structure containing an array of (named) arguments etc.
  [out]     VARIANT    *pVarResult, //Pointer to the location where the result is to be stored, or NULL (DISPATCH_PROPERTYPUT... ignored)
  [out]     EXCEPINFO  *pExcepInfo, //Pointer to a structure that contains exception information
  [out]     UINT       *puArgErr //The index within rgvarg of the first argument that has an error (??)
);
```

---

Dispatch function DispInvoke that provides a standard implementation of Invoke:

**DispInvoke** function (oleauto.h):

https://learn.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-dispinvoke

```cpp
HRESULT DispInvoke(
  void       *_this, //An implementation of the IDispatch interface described by ptinfo.
  ITypeInfo  *ptinfo, //The type information that describes the interface.
  DISPID     dispidMember, //The member to be invoked. Use GetIDsOfNames or the object's documentation to obtain the DISPID.
  WORD       wFlags, //Flags describing the context of the Invoke call.
  DISPPARAMS *pparams, //Pointer to a structure containing an array of arguments, 
                       //an array of argument DISPIDs for named arguments, and counts for number of elements in the arrays.
  VARIANT    *pvarResult, //Pointer to where the result is to be stored, or Null if the caller expects no result.
                          //This argument is ignored if DISPATCH_PROPERTYPUT or DISPATCH_PROPERTYPUTREF is specified.
  EXCEPINFO  *pexcepinfo, //Pointer to a structure containing exception information. 
                          //This structure should be filled in if DISP_E_EXCEPTION is returned.
  UINT       *puArgErr //The index within rgvarg of the first argument that has an error.
                       //Arguments are stored in pdispparams->rgvarg in reverse order, 
                       //so the first argument is the one with the highest index in the array. 
                       //This parameter is returned only when the resulting return value is DISP_E_TYPEMISMATCH or DISP_E_PARAMNOTFOUND.
);
```

**wFlags:**

**DISPATCH_METHOD**: The member is invoked as a method. If a property has the same name, both this and the DISPATCH_PROPERTYGET flag can be set.

**DISPATCH_PROPERTYGET** The member is retrieved as a property or data member.

**DISPATCH_PROPERTYPUT** The member is changed as a property or data member.

**DISPATCH_PROPERTYPUTREF** The member is changed by a reference assignment, rather than a value assignment. This flag is valid only when the property accepts a reference to an object. 

---

**JNI Export function that calls pIDispatch->Invoke (Dispatch.cpp)**

```cpp
//JNI *wrapper around windows call
  JNIEXPORT jobject JNICALL Java_com_jacob_com_Dispatch_invokev(JNIEnv *env,
                                                                jclass clazz,
                                                                jobject disp, 
                                                                jstring name, 
                                                                jint dispid,
                                                                jint lcid, 
                                                                jint wFlags, 
                                                                jobjectArray vArg, 
                                                                jintArray uArgErr)
```

**Direct call from Jacob JNI (Dispatch.cpp)**

```cpp
//The call itself
pIDispatch->Invoke(dispID, 
                   IID_NULL,
                   lcid, 
                   (WORD)wFlags, 
                   &dispparams, 
                   v, 
                   &excepInfo, 
                   (unsigned int *)uAE);
```

---

**Java native Invoke**
```java
native Variant invokev(Dispatch dispatchTarget, 
	                   String name, 
	                   int dispID, 
	                   int lcid, 
	                   int wFlags, 
	                   Variant[] vArg, 
	                   int[] uArgErr)
```

**Invokev (calls Invoke native)**
```java
Variant invokev(Dispatch dispatchTarget, String name, int wFlags, Variant[] vArg, int[] uArgErr)
        invokev(dispatchTarget, name, 0, Dispatch.LOCALE_SYSTEM_DEFAULT, wFlags, vArg, uArgErr)

Variant invokev(Dispatch dispatchTarget, int dispID,  int wFlags, Variant[] vArg, int[] uArgErr)
        invokev(dispatchTarget, null, dispID, Dispatch.LOCALE_SYSTEM_DEFAULT, wFlags, vArg, uArgErr)
//not implemented Variant invokev(Dispatch dispatchTarget, String name, int wFlags, Variant[] vArg, int[] uArgErr, int wFlagsEx) 
//not implemented invokev(dispatchTarget, name, 0, Dispatch.LOCALE_SYSTEM_DEFAULT, wFlags, vArg, uArgErr) 
```
Conclusion: non-native Invokev provide native Invokev with lcid = Dispatch.LOCALE_SYSTEM_DEFAULT, and depending on 
String or int in second parameter chooses (name, 0) vs. (null dispID), rest argument passed as-is.

**Invoke**

```java
Variant invoke(Dispatch dispatchTarget, String name, int dispID, int lcid, int wFlags,                           Object[] oArg, int[] uArgErr)
       invokev(dispatchTarget,                 name,     dispID,     lcid,     wFlags, VariantUtilities.objectsToVariants(oArg),      uArgErr)

Variant invoke(Dispatch dispatchTarget, String name, int wFlags, Object[] oArg, int[] uArgErr)
       invokev(dispatchTarget, name, wFlags, VariantUtilities.objectsToVariants(oArg), uArgErr)

Variant invoke(Dispatch dispatchTarget, int dispID,	int wFlags, Object[] oArg, int[] uArgErr)
       invokev(dispatchTarget, dispID, wFlags, VariantUtilities.objectsToVariants(oArg), uArgErr)
```

Conclusion: Main feature of Invoke is converting Object[] oArg to Variant[] vArg with **VariantUtilities.objectsToVariants**

One signature is one-to-one call for native invokev, other two are the same 5 arguments choosing between name and dispID

**CallN**

CallN is only called by Call

```java
Variant callN(Dispatch dispatchTarget, String name,	Object... args)
      invokev(dispatchTarget, name, Dispatch.Method | Dispatch.Get, VariantUtilities.objectsToVariants(args), new int[args.length])

Variant callN(Dispatch dispatchTarget, int dispID,	Object... args)
      invokev(dispatchTarget, dispID, Dispatch.Method | Dispatch.Get, VariantUtilities.objectsToVariants(args), new int[args.length])
```

Conclusion: calls CallN calls 5-arguments invokev, converts Object... args with **VariantUtilities.objectsToVariants**,
and most importantly provides Dispatch.Method | Dispatch.Get wFlags
(If a property has the same name, both this and the DISPATCH_PROPERTYGET flag can be set)


**Call**

```java
Variant call(Dispatch dispatchTarget, String name)
callN(dispatchTarget, name, NO_VARIANT_ARGS)

Variant call(Dispatch dispatchTarget, int dispid)
callN(dispatchTarget, dispid, NO_VARIANT_ARGS)

Variant call(Dispatch dispatchTarget, String name, Object... attributes)
callN(dispatchTarget, name, attributes)

Variant call(Dispatch dispatchTarget, int dispid, Object... attributes)
callN(dispatchTarget, dispid, attributes)
```

**Put**

```java
void put(Dispatch dispatchTarget, String name, Object val) 
void put(Dispatch dispatchTarget, int dispid,  Object val) 
```

Calls this:

```java
invoke(dispatchTarget, name,   Dispatch.Put, new Object[] { val }, new int[1])
invoke(dispatchTarget, dispid, Dispatch.Put, new Object[] { val }, new int[1])
```

**Get**

```java
Variant get(Dispatch dispatchTarget, String name)
Variant get(Dispatch dispatchTarget, int dispid) 
```

Calls this:

```java
invokev(dispatchTarget, name,   Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS)
invokev(dispatchTarget, dispid, Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS)
```





**Invokev** -->> invokev__native

**Invoke** -->> Invokev 1 

**CallN** -->> Invokev 1

**Call** -->> CallN 2

**Put** -->> Invoke 3

**Get** -->> Invokev 1

---

Invokev called by: Invoke (only by Put), CallN (only by Call), Get (calls invokev directly)

CallN called by: Call

Invoke called by: Put

---

---

Get - Variant Dispatch dispatchTarget, String name -->> Invokev

Get = Invokev(dispatchTarget, **name**, Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS) 


Put - Void Dispatch dispatchTarget, String name, Object val -->> Invoke: -->> Invokev

Invoke: 1> dispatchTarget, 2> name, 3> Dispatch.Put, 4> new Object[] { val } (**single object oArg**), 5> new int[1] (**standard uArgErr**)

Invokev: 1> dispatchTarget, 2> name, 3> 0, 4> Dispatch.LOCALE_SYSTEM_DEFAULT, 5> wFlags, 6> vArg, 7> uArgErr



Put = 


Call -->> CallN --> Invokev

---




```Java
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class ExcelDispatchTest {
	public static void main(String[] args) {
		ComThread.InitSTA();

		ActiveXComponent xl = new ActiveXComponent("Excel.Application");
		try {
			System.out.println("version=" + xl.getProperty("Version"));
			System.out.println("version=" + Dispatch.get(xl, "Version"));
			Dispatch.put(xl, "Visible", new Variant(true));
			Dispatch workbooks = xl.getProperty("Workbooks").toDispatch();
			Dispatch workbook = Dispatch.get(workbooks, "Add").toDispatch();
			Dispatch sheet = Dispatch.get(workbook, "ActiveSheet").toDispatch();
			Dispatch a1 = Dispatch.invoke(sheet, "Range", Dispatch.Get,
					new Object[] { "A1" }, new int[1]).toDispatch();
			Dispatch a2 = Dispatch.invoke(sheet, "Range", Dispatch.Get,
					new Object[] { "A2" }, new int[1]).toDispatch();
			Dispatch.put(a1, "Value", "123.456");
			Dispatch.put(a2, "Formula", "=A1*2");
			System.out.println("a1 from excel:" + Dispatch.get(a1, "Value"));
			System.out.println("a2 from excel:" + Dispatch.get(a2, "Value"));
			Variant f = new Variant(false);
			Dispatch.call(workbook, "Close", f);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			xl.invoke("Quit", new Variant[] {});
			ComThread.Release();
		}
	}

}

```


public static final ints:

> LOCALE_SYSTEM_DEFAULT = 2048; // Used to set the locale in a call. The user locale is another option

> Method = 1; /** used by callN() and callSubN()

> Get = 2; /** used by callN() and callSubN()

> Put = 4; //** used by put()

> int PutRef = 8; //** not used, probably intended for putRef() 

> private final static Object[] NO_OBJECT_ARGS = new Object[0];

> private final static Variant[] NO_VARIANT_ARGS = new Variant[0];

> private final static int[] NO_INT_ARGS = new int[0];


## Dispatch methods:

Cover for call to underlying invokev():

### Get

**Examples**

```java
Dispatch.get(xl, "Version");

Dispatch.get(workbooks, "Add").toDispatch();

Dispatch.get(workbook, "ActiveSheet").toDispatch();

Dispatch.get(a1, "Value");

Dispatch.get(a2, "Value");
```

- Dispatch dispatchTarget,  **String name**:

```java
invokev(dispatchTarget, name, Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS);
```

- Dispatch dispatchTarget, **int dispid**

```java
invokev(dispatchTarget, dispid, Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS);
```

### Put

**Examples**

```java
Dispatch.put(xl, "Visible", new Variant(true));

Dispatch.put(a1, "Value", "123.456");

Dispatch.put(a2, "Formula", "=A1\*2");
```

- Dispatch dispatchTarget, **String name**, Object val

```java
invoke(dispatchTarget, name, Dispatch.Put, new Object[] { val }, new int[1]);
```

- Dispatch dispatchTarget, **int dispid**, Object val

```java
invoke(dispatchTarget, dispid, Dispatch.Put, new Object[] { val }, new int[1]);
```

### Call

**Examples**

```java
Dispatch.call(workbook, "Close", f);
```


- Dispatch dispatchTarget, java.lang.String name

```java
callN(dispatchTarget, name, NO_VARIANT_ARGS);
```

- Dispatch dispatchTarget, java.lang.String name, java.lang.Object... attributes

```java
callN(dispatchTarget, name, attributes);
```

- Dispatch dispatchTarget, int dispid

```java
callN(dispatchTarget, dispid, NO_VARIANT_ARGS);
```

- Dispatch dispatchTarget, int dispid, java.lang.Object... attributes

```java
callN(dispatchTarget, dispid, attributes);
```

### Invoke

**Examples**

```java
Dispatch a1 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] { "A1" }, new int[1]).toDispatch();

Dispatch a2 = Dispatch.invoke(sheet, "Range", Dispatch.Get, new Object[] { "A2" }, new int[1]).toDispatch();

xl.invoke("Quit", new Variant[] {});
```

- Dispatch dispatchTarget, java.lang.String name, int wFlags, java.lang.Object[] oArg, int[] uArgErr

```java
invokev(dispatchTarget, name, wFlags, VariantUtilities.objectsToVariants(oArg), uArgErr);
```

- Dispatch dispatchTarget, java.lang.String name, int dispID, int lcid, int wFlags, java.lang.Object[] oArg, int[] uArgErr

```java
invokev(dispatchTarget, name, dispID, lcid, wFlags,	VariantUtilities.objectsToVariants(oArg), uArgErr);
```

- Dispatch dispatchTarget, int dispID, int wFlags, java.lang.Object[] oArg, int[] uArgErr)

```java
invokev(dispatchTarget, dispID, wFlags, VariantUtilities.objectsToVariants(oArg), uArgErr);
```
