# 把它发布到 GitHub —— 手把手

给第一次开源的人写的。每一步都说明**为什么**，不只是敲什么。

---

## 先分清两个东西

| | 是什么 | 在哪 |
|---|---|---|
| **git** | 一个版本管理程序，记录你文件夹里每次改了什么 | 你电脑上，已经装好了（git 2.50.1） |
| **GitHub** | 一个网站，把 git 仓库放在网上给别人看 | github.com |

`git` 管本地，GitHub 管线上。你要做的是：**先在本地建仓库，再推上去。**

---

## 需要四样东西

1. **gh** —— GitHub 官方命令行工具（正在后台安装）
2. **你的身份** —— git 得知道每次提交是谁写的（姓名 + 邮箱）
3. **登录** —— 让 gh 拿到你账号的授权
4. **仓库名** —— 就叫 `pptx-slim`

---

## 第 1 步：装 gh

```bash
brew install gh
gh --version
```

**为什么不用网页上传？** 网页上传不支持文件夹结构，也记不住历史。命令行一次配好，以后一条命令就能更新。

---

## 第 2 步：告诉 git 你是谁

```bash
git config --global user.name "你的名字"
git config --global user.email "你的邮箱"
```

**为什么？** 每次提交（commit）都会盖上这个名字和时间戳，写在历史里。别人看你的仓库，看到的就是这个。

**邮箱填哪个？** 用你 GitHub 账号绑定的邮箱，这样提交能关联到你的头像。**不想公开邮箱**的话，GitHub 提供了一个隐私邮箱，形如 `12345678+用户名@users.noreply.github.com`，在 GitHub → Settings → Emails 里能找到。用它更安全。

验证：

```bash
git config --global user.name
git config --global user.email
```

---

## 第 3 步：登录 GitHub

```bash
gh auth login
```

它会依次问你：

```
? What account do you want to log into?        → GitHub.com
? What is your preferred protocol for Git operations?  → HTTPS
? Authenticate Git with your GitHub credentials?       → Yes
? How would you like to authenticate?          → Login with a web browser
```

然后它给你一个**一次性代码**，形如 `A1B2-C3D4`，并自动打开浏览器。把代码粘进去，点授权，就完事了。

验证：

```bash
gh auth status
```

看到 `Logged in to github.com as 你的用户名` 就成了。

**它在做什么？** gh 拿到一个 token（令牌）存在 macOS 钥匙串里。以后 git 推代码，系统自动拿它去认证——**你不用输密码**。

---

## 第 4 步：建仓库并推上去

仓库已经在 `~/pptx-slim` 准备好了。先做个本地提交：

```bash
cd ~/pptx-slim
git init
git add -A
git commit -m "Initial release: pptx-slim 1.0.0"
```

然后一句话在 GitHub 上建仓库、关联、推送：

```bash
gh repo create pptx-slim --public --source=. --push
```

它会问你确认，回车即可。

**这条命令干了三件事：**
- 在 GitHub 上建一个叫 `pptx-slim` 的公开仓库
- 把你本地这个仓库设成它的 `origin`（远程地址）
- 把所有提交推上去

---

## 第 5 步：确认

```bash
gh repo view --web
```

浏览器打开你的仓库。看到 README 渲染出来了，就成了。

---

## 以后怎么更新

改完代码，走这三步：

```bash
cd ~/pptx-slim
git add -A
git commit -m "说清楚这次改了什么"
git push
```

`git add` 挑要记录的改动，`git commit` 拍一张快照并写说明，`git push` 推到 GitHub。

看当前状态：

```bash
git status     # 哪些文件改了还没提交
git log --oneline -10   # 最近 10 次提交
```

---

## 常见报错

| 报错 | 原因 | 怎么办 |
|---|---|---|
| `Author identity unknown` | 没做第 2 步 | `git config --global user.name/user.email` |
| `remote origin already exists` | 已经关联过了 | `git remote set-url origin <地址>`，或直接 `git push` |
| `Authentication failed` | 没登录 / token 过期 | `gh auth login` 重来一遍 |
| `Updates were rejected` | GitHub 上有你本地没有的提交 | `git pull --rebase` 再 `git push` |
| `Permission denied (publickey)` | 用了 SSH 但没配密钥 | 重新 `gh auth login`，协议选 **HTTPS** |

---

## 关于协议：HTTPS 还是 SSH？

选 **HTTPS**。配合 gh，令牌自动管理，零配置。

SSH 需要自己生成密钥、贴到 GitHub、配 config——多三步，初学者没必要。

---

## 仓库里有些什么

```
pptx-slim/
├── README.md              # 英文说明（外国人先看到这个）
├── README.zh-CN.md        # 中文说明
├── LICENSE                # MIT 许可证
├── pyproject.toml         # 让别人能 pip install
├── .gitignore             # 告诉 git 哪些文件不要上传
├── pptx_slim.py           # 引擎本体（单文件，无依赖除了 Pillow）
└── macos/
    ├── droplet.applescript  # 拖拽 App 的源码
    └── build.sh             # 一键编译 App
```

**`.gitignore` 里为什么忽略 `*.pptx`？** 那是你的真实汇报文件，不该传到公开仓库——里面可能有你的实验数据。别人 clone 下来也不会被一堆大文件拖慢。

---

## 开源前检查清单

- [ ] 仓库里**没有任何**私人文件（汇报、实验数据、截图）
- [ ] README 里的 `USERNAME` 全部换成了你的 GitHub 用户名
- [ ] `pyproject.toml` 里的 `Homepage` / `Issues` 链接换好了
- [ ] `LICENSE` 里的版权人写的是你
- [ ] 用 `git status` 确认没有意外文件混进来

---

## 一个提醒：许可证

选了 **MIT**。意思很简单：**谁都可以免费用、改、甚至拿去卖，但必须保留你的版权声明，且你不能担保它一定好用。**

对一个小工具来说这是最省事、对使用者最友好的选择。想更严格（比如要求别人改了也必须开源）就用 GPL-3.0，但那样用的人会少一些。
