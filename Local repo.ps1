## SETUP LOCAL GIT REPO WITH A LOCAL REMOTE
# the main elements:
# - remote repo must be initialized with --bare parameter
# - local repo must be initialized
# - local repo must have at least one commit that properly initializes a branch(root of the commit tree)
# - local repo needs to have a remote
# - local repo branch must have an upstream branch on the remote

{ # the brackets are optional, they allow to copy paste into terminal and run entire thing without interruptions, run without them to see which cmd outputs what

  Set-Location ~
  Remove-Item -rf ~/test_git_local_repo/

  ## Option A - clean slate - you have nothing yet

  mkdir -p ~/test_git_local_repo/option_a ; cd ~/test_git_local_repo/option_a
  git init --bare local_remote.git # first setup the local remote
  git clone local_remote.git local_repo # creates a local repo in dir local_repo
  cd ~/test_git_local_repo/option_a/local_repo
  git remote -v show origin # see that git clone has configured the tracking
  touch README.md ; git add . ; git commit -m "initial commit on master" # properly init master
  git push origin master # now have a fully functional setup, -u not needed, git clone does this for you

  # check all is set-up correctly
  git pull # check you can pull
  git branch -avv # see local branches and their respective remote upstream branches with the initial commit
  git remote -v show origin # see all branches are set to pull and push to remote
  git log --oneline --graph --decorate --all # see all commits and branches tips point to the same commits for both local and remote

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

  ## Option C - you already have a directory with some files and you want it to be a git repo with a local remote

  mkdir -p ~/test_git_local_repo/option_c ; cd ~/test_git_local_repo/option_c
  git init --bare local_remote.git # first setup the local remote

  # simulate a pre-existing directory with some files
  mkdir local_repo ; cd local_repo ; touch README.md fileB

  # make a pre-existing directory a git repo and connect it with local remote
  cd ~/test_git_local_repo/option_c/local_repo
  git init
  git add . ; git commit -m "inital commit on master" # properly init master
  git remote add origin ~/test_git_local_repo/option_c/local_remote.git
  git remote -v show origin # see there is no the tracking configured (unlike with git clone), so you need to push with -u
  git push -u origin master # -u to set upstream

  # check all is set-up correctly
  git pull # check you can pull
  git branch -avv # see local branch and its remote upstream with the initial commit
  git remote -v show origin # see all remote branches are set to pull and push to remote
  git log --oneline --graph --decorate --all # see all commits and branches tips point to the same commits for both local and remote
}
