#!/bin/bash
# 收集社区热帖：Twitter 热榜 + 小红书推荐流 + Reddit LocalLLaMA
# 输出到 /tmp/social_posts.md，供 generate_docx.py --x-posts-file 使用

export PATH="$HOME/.hermes/node/bin:$PATH"

DATE=$(date +%Y-%m-%d)
OUT=/tmp/social_posts.md

echo "# 社区热帖精选 · $DATE" > "$OUT"
echo "" >> "$OUT"

# ── Twitter/X 热榜 ──────────────────────────────────────────
echo "## X / Twitter 热榜" >> "$OUT"
echo "" >> "$OUT"

TWITTER=$(opencli twitter trending 2>/dev/null)
if [ -n "$TWITTER" ]; then
    echo "$TWITTER" | awk '
    /^- rank:/ { rank=$3 }
    /^  topic:/ {
        topic = substr($0, index($0,$2))
        gsub(/^[ \t]+/, "", topic)
        printf "%s. %s\n", rank, topic
    }
    /^  category:/ {
        cat = substr($0, index($0,$2))
        gsub(/^[ \t]+/, "", cat)
        printf "   （%s）\n", cat
    }
    ' | head -30 >> "$OUT"
else
    echo "（获取失败）" >> "$OUT"
fi
echo "" >> "$OUT"

# ── 小红书推荐流 ─────────────────────────────────────────────
echo "## 小红书推荐流" >> "$OUT"
echo "" >> "$OUT"

XHS=$(opencli xiaohongshu feed --limit 10 2>/dev/null)
if [ -n "$XHS" ]; then
    echo "$XHS" | awk '
    BEGIN { i=0 }
    /^- id:/ { i++ }
    /^  title:/ {
        title = substr($0, index($0,$2))
        gsub(/^[ \t]+/, "", title)
        printf "%s. %s\n", i, title
    }
    /^  author:/ {
        author = substr($0, index($0,$2))
        gsub(/^[ \t]+/, "", author)
        printf "   作者：%s\n", author
    }
    /^  likes:/ {
        likes = substr($0, index($0,$2))
        gsub(/^[ \t]+|'"'"'/, "", likes)
        printf "   点赞：%s\n", likes
    }
    /^  url:/ {
        url = substr($0, index($0,$2))
        gsub(/^[ \t]+/, "", url)
        printf "   链接：%s\n", url
    }
    ' >> "$OUT"
else
    echo "（获取失败）" >> "$OUT"
fi
echo "" >> "$OUT"

# ── Reddit LocalLLaMA 热帖 ───────────────────────────────────
echo "## Reddit r/LocalLLaMA 热帖" >> "$OUT"
echo "" >> "$OUT"

REDDIT=$(opencli reddit hot --subreddit LocalLLaMA --limit 5 2>/dev/null)
if [ -n "$REDDIT" ]; then
    echo "$REDDIT" | awk '
    /^- rank:/ { rank=$3; in_title=0 }
    /^  title:/ {
        if ($2 == ">-") {
            in_title=1
        } else {
            title = substr($0, index($0,$2))
            gsub(/^[ \t]+|^'"'"'|'"'"'$/, "", title)
            printf "%s. %s\n", rank, title
            in_title=0
        }
    }
    in_title && /^    / {
        title = substr($0, 5)
        gsub(/^[ \t]+|[ \t]+$/, "", title)
        printf "%s. %s\n", rank, title
        in_title=0
    }
    /^  score:/ {
        score=$2
        printf "   赞数：%s\n", score
    }
    /^  url:/ {
        url=$2
        printf "   链接：%s\n\n", url
    }
    ' >> "$OUT"
else
    echo "（获取失败）" >> "$OUT"
fi

echo "✅ 社区热帖已收集 → $OUT"
