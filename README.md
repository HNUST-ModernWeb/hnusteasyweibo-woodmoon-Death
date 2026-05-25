[![Review Assignment Due Date](https://classroom.github.com/assets/deadline-readme-button-22041afd0340ce965d47ae6ef1cefeee28c7c493a6346c4f15d667ab976d596c.svg)](https://classroom.github.com/a/QLSGUFuT)
[![Open in Visual Studio Code](https://classroom.github.com/assets/open-in-vscode-2e0aaae1b6195c2367325f4f02e2d04e9abb55f0b24a779b69b11b9e10269abc.svg)](https://classroom.github.com/online_ide?assignment_repo_id=23340050&assignment_repo_type=AssignmentRepo)

# 心语栈 社交分享平台

这是一个直接进入最终阶段的前后端分离社交分享平台项目。

## 技术栈

- 前端：Vue3、Vite、Vue Router、Pinia、Axios
- 后端：Java 8、Spring Boot 2.7、Spring MVC、MyBatis、MySQL
- 功能：注册登录、信息流、发布文字、图片上传、评论、点赞、个人主页、管理员管理、权限控制

## 目录结构

```text
webwork/
  frontend/   Vue3 前端
  backend/    Spring Boot 后端
  方案.md
  Agent.md
```

## 数据库初始化

1. 创建 MySQL 数据库和表：

```sql
source backend/src/main/resources/schema.sql;
```

2. 修改后端数据库配置：

```text
backend/src/main/resources/application.yml
```

默认配置：

```yaml
spring:
  datasource:
    url: jdbc:mysql://localhost:3306/social_share?useUnicode=true&characterEncoding=utf8&useSSL=false&serverTimezone=Asia/Shanghai
    username: root
    password: root
```

## 启动后端

项目已提供便携版 Maven 启动脚本：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/start-backend.ps1
```

也可以手动进入后端目录：

```bash
cd backend
..\tools\apache-maven-3.9.9\bin\mvn.cmd spring-boot:run
```

后端默认端口：

```text
http://localhost:8080
```

## 启动前端

项目已提供前端启动脚本：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/start-frontend.ps1
```

也可以手动进入前端目录：

```bash
cd frontend
npm install
npm run dev
```

前端默认端口：

```text
http://localhost:5173
```

Vite 已配置代理：

- `/api` -> `http://localhost:8080`
- `/uploads` -> `http://localhost:8080`

## 核心页面

- `/` 信息流
- `/login` 登录
- `/register` 注册
- `/publish` 发布
- `/posts/:id` 分享详情
- `/users/:id` 个人主页
- `/admin` 管理页

## 权限说明

- 游客只能浏览公开内容。
- 普通用户可以发布、点赞、评论、删除自己的内容。
- 管理员可以删除、隐藏、恢复任意内容。
- 后端负责最终权限校验，前端隐藏按钮只用于改善体验。

## 注意事项

- 当前 token 是轻量 Base64 演示实现，课程项目可用；生产环境应替换为 JWT 或服务端 Session。
- 密码使用 BCrypt 哈希存储。
- 上传图片保存在 `backend/uploads/`，通过 `/uploads/**` 访问。
- 管理员账号需要直接在数据库中将 `users.role` 设置为 `ADMIN`。
- MySQL 已检测到本机服务名为 `MySQL80`，客户端路径为 `C:\Program Files\MySQL\MySQL Server 8.0\bin\mysql.exe`。
- 可用 `powershell -ExecutionPolicy Bypass -File scripts/init-db.ps1` 初始化数据库。
