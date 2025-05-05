# Samples from Jacob project


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

### Call hierarchy

**CallN**

Variant callN(Dispatch dispatchTarget, String name,	Object... args)
Variant callN(Dispatch dispatchTarget, int dispID,	Object... args)

Calls this:

invokev(dispatchTarget, name,   Dispatch.Method | Dispatch.Get, VariantUtilities.objectsToVariants(args), new int[args.length])
invokev(dispatchTarget, dispID, Dispatch.Method | Dispatch.Get,	VariantUtilities.objectsToVariants(args), new int[args.length])

**Invoke**

Variant invoke(Dispatch dispatchTarget, String name, int dispID, int lcid, int wFlags, Object[] oArg, int[] uArgErr)
Variant invoke(Dispatch dispatchTarget, String name, int wFlags, Object[] oArg, int[] uArgErr)
Variant invoke(Dispatch dispatchTarget, int dispID,	int wFlags, Object[] oArg, int[] uArgErr)

Calls this:

invokev(dispatchTarget, name, dispID, lcid, wFlags,	VariantUtilities.objectsToVariants(oArg), uArgErr)
invokev(dispatchTarget, name, wFlags, VariantUtilities.objectsToVariants(oArg), uArgErr)
invokev(dispatchTarget, dispID, wFlags, VariantUtilities.objectsToVariants(oArg), uArgErr)

**Call**

Variant call(Dispatch dispatchTarget, String name)
Variant call(Dispatch dispatchTarget, String name, Object... attributes)
Variant call(Dispatch dispatchTarget, int dispid) 
Variant call(Dispatch dispatchTarget, int dispid, Object... attributes)


Calls this:

callN(dispatchTarget, name, NO_VARIANT_ARGS)
callN(dispatchTarget, name, attributes)
callN(dispatchTarget, dispid, NO_VARIANT_ARGS)
callN(dispatchTarget, dispid, attributes)


**Put**

void put(Dispatch dispatchTarget, String name, Object val) 
void put(Dispatch dispatchTarget, int dispid,  Object val) 

Calls this:

invoke(dispatchTarget, name,   Dispatch.Put, new Object[] { val }, new int[1])
invoke(dispatchTarget, dispid, Dispatch.Put, new Object[] { val }, new int[1])


**Invokev (Invoke native)**

native Variant invokev(Dispatch dispatchTarget, String name, int dispID, int lcid, int wFlags, Variant[] vArg, int[] uArgErr)
Variant invokev(Dispatch dispatchTarget, String name, int wFlags, Variant[] vArg, int[] uArgErr)
Variant invokev(Dispatch dispatchTarget, String name, int wFlags, Variant[] vArg, int[] uArgErr, int wFlagsEx) 
Variant invokev(Dispatch dispatchTarget, int dispID,  int wFlags, Variant[] vArg, int[] uArgErr)


Calls this:

First is native invokev!
invokev(dispatchTarget, name, 0, Dispatch.LOCALE_SYSTEM_DEFAULT, wFlags, vArg, uArgErr)
invokev(dispatchTarget, name, 0, Dispatch.LOCALE_SYSTEM_DEFAULT, wFlags, vArg, uArgErr)
invokev(dispatchTarget, null, dispID, Dispatch.LOCALE_SYSTEM_DEFAULT, wFlags, vArg, uArgErr)

**Get**

Variant get(Dispatch dispatchTarget, String name)
Variant get(Dispatch dispatchTarget, int dispid) 

Calls this:

invokev(dispatchTarget, name,   Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS)
invokev(dispatchTarget, dispid, Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS)












