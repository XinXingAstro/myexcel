# !/bin/bash

# change to current dir
cd `dirname $0`
echo "当前目录：`pwd`"
git checkout dev
# get sha
read -p "请输入上线SHA码: " sha
# confirm sha
echo -e "确认重置dev分支至基线号: \033[1;4;31m[$sha]\033[0m y/n : "
read confirm
if [[ "$confirm" == "y" || "$confirm" == "yes" || "$confirm" == "" ]]; then
    echo "开始查找sha码"
    # check sha in commit tree
    commitNum=$(git rev-list HEAD --count)
    history=$(git log --pretty=format:"%H" -$commitNum)
    shaArray=($history)
    for ((i = 0; i < $commitNum; i++)); do
        echo ${shaArray[i]}
        if [ "$sha" == "${shaArray[i]}" ]; then
            break
        fi  
    done
    if [ $i -eq $commitNum ]; then
        echo "未找到相同sha码，sha码输入错误！"
        read -p "请按任意键退出"
        exit
    else
        read -p "已找到指定sha，确认是否重置 y/n : " confirmReset
        if [[ "$confirmReset" == "y" || "$confirmReset" == "yes" || "$confirmReset" == "" ]]; then
            # reset dev to sha
            echo "开始重置dev分支" 
            git checkout dev
            git pull
            git reset $sha
            git push -f
            # check whether success
            curSha=$(git log --pretty=format:"%H" -1)
            if [ "$sha" == "$curSha" ]; then 
                echo "dev分支重置成功！"
            else 
                echo "dev分支重置失败！"
            fi
        fi
    fi
fi

read -p "请按任意键退出"
exit
