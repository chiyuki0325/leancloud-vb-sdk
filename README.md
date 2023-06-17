# ☁️ [WIP] LeanCloud VB SDK

### 依赖

[VBJSON](https://github.com/YidaozhanYa/VBJSON)

## 初始化

```vb
Dim LC As New cLeanCloud
LC.Initialize "YourAppId", "YourAppKey", "AppHostDomain"
```

### 数据存储

已实现的功能如下：

#### 对象

```vb
Dim Obj As LCObject
Set Obj = LC.Object("Hello")
Obj!IntValue = 123
Obj.Save
    
MsgBox Obj.ObjectID  '6485754238ad8b7c0f463fd5
```

#### 查询

```vb
Dim Query As LCQuery, Obj As LCObject
Set Query = LC.Query("Hello")
Set Obj = Query.GetObject("6485754238ad8b7c0f463fd5")
MsgBox Obj!IntValue  '123
```