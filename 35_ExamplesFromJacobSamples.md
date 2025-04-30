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

private final static Object[] NO_OBJECT_ARGS = new Object[0];
private final static Variant[] NO_VARIANT_ARGS = new Variant[0];
private final static int[] NO_INT_ARGS = new int[0];


## Dispatch methods:

Cover for call to underlying invokev():

### Get

```text
Dispatch.get(xl, "Version");

Dispatch.get(workbooks, "Add").toDispatch();

Dispatch.get(workbook, "ActiveSheet").toDispatch();

Dispatch.get(a1, "Value"));

Dispatch.get(a2, "Value"));
```

Dispatch dispatchTarget,  **String name**:

```java
invokev(dispatchTarget, name, Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS);
```

Dispatch dispatchTarget, **int dispid**

```java
invokev(dispatchTarget, dispid, Dispatch.Get, NO_VARIANT_ARGS, NO_INT_ARGS);
```

### Put

Dispatch.put(xl, "Visible", new Variant(true));

Dispatch.put(a1, "Value", "123.456");

Dispatch.put(a2, "Formula", "=A1\*2");


.put:  Dispatch dispatchTarget, **String name**, Object val

invoke(dispatchTarget, name, Dispatch.Put, new Object[] { val }, new int[1]);

.put:  Dispatch dispatchTarget, **int dispid**, Object val

invoke(dispatchTarget, dispid, Dispatch.Put, new Object[] { val }, new int[1]);

### Call

```java
Dispatch.call(workbook, "Close", f);
```

.call   (Dispatch dispatchTarget, int dispid) 

callN(dispatchTarget, dispid, NO_VARIANT_ARGS);


.call   (Dispatch dispatchTarget, int dispid, java.lang.Object... attributes) 

callN(dispatchTarget, dispid, attributes);


.call   (Dispatch dispatchTarget, java.lang.String name) 

callN(dispatchTarget, name, NO_VARIANT_ARGS);


.call   (Dispatch dispatchTarget, java.lang.String name, java.lang.Object... attributes) 

callN(dispatchTarget, name, attributes);






















