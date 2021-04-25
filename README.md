# 漂流书架

---

>   关于我，欢迎关注：
>
>   Github主页：[MrCai-starter (github.com)](https://github.com/MrCai-starter)
>
>   个人邮箱：1014305148@qq.com



## 项目介绍

本项目为电子科技大学“互联网+”学生联合协会的“漂流书架”活动开发。

实现了最简单的HCI（人机交互）、Excel自动化办公、浏览器模拟操作，为协会人员在后台登记书籍信息、捐赠信息提供简单接口。

目前版本为v1.1，拥有：

①一整行的输入与在Excel内的持久化存储；

②任意节点退出；

③模拟浏览器自动生成、下载二维码。

三个功能。



## 使用方法

1.  根据提示依次输入数据，或进行选择(y/n)。
2.  在任意一点均可输入“quit”指令以退出。
3.  在浏览器的“批量模板”页面，先点击“书籍信息”的“批量生码”，再点击“捐赠证明”的“批量生码”。



## 注意事项

1.  文件“category.json”与“漂流书架.xlsx”需放在同一目录下，否则将无法运行。
2.  模拟浏览器中，除了两个“批量生码”按钮，其他什么都不要点！！！



## TODO

1.  中途使用“quit”指令退出时：如果用户输入“n”反悔，则返回刚才的输入项，而非下一个输入项。
2.  添加“撤销该行操作”的接口。
3.  对“autofit”部分进行修改，美化表格展示效果。