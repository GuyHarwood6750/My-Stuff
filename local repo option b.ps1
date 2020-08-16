## Option B - you already have a local git repo and you want to connect it to a local remote
mkdir -p ~/test_git_local_repo/option_b ; cd ~/test_git_local_repo/option_b
git init --bare local_remote.git # first setup the local remote

# simulate a pre-existing git local repo you want to connect with the local remote
mkdir local_repo ; cd local_repo
git init # if not yet a git repo
touch README.md ; git add . ; git commit -m "initial commit on master" # properly init master
git checkout -b develop ; touch fileB ; git add . ; git commit -m "add fileB on develop" # create develop and fake change

# connect with local remote
cd ~/test_git_local_repo/option_b/local_repo
git remote add origin ~/test_git_local_repo/option_b/local_remote.git
git remote -v show origin # at this point you can see that there is no the tracking configured (unlike with git clone), so you need to push with -u
git push -u origin master # -u to set upstream
git push -u origin develop # -u to set upstream; need to run this for every other branch you already have in the project

# check all is set-up correctly
git pull # check you can pull
git branch -avv # see local branch(es) and its remote upstream with the initial commit
git remote -v show origin # see all remote branches are set to pull and push to remote
git log --oneline --graph --decorate --all # see all commits and branches tips point to the same commits for both local and remote
