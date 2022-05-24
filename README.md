# py_docx_ripper
A Python script that traverses recursively down a set directory starting point to parse Microsoft Word .docx contents.

This current version of the readme.md outlines the setup necessary to create a virtual environment that runs jupyter in order to assist in the process of discovery of the structure and contents of the .docx XML content payloads.

# Environment Setup & Config
R&D done on MacOS Big Sur - setup will vary for *nix / win environments.

## Install / Configure Pyenv
1 - Install pyenv for python environment version management. Then set pyenv paths to $PATH for your env

```zsh
brew install pyenv
export PYENV_ROOT=$(pyenv root)
export PATH="$PYENV_ROOT/shims:$PATH"
```
2 - Install the virtual environment manager of pyenv to enable multiple python version management for future dev

```zsh
brew install pyenv-virtualenv
eval "$(pyenv virtualenv-init -)"
```

3 - (OPTIONAL) set local / global python versions as necessary

```zsh
pyenv local 3.10.2
```

## Install / Configure Pipenv
We will use pipenv to mantain and manage the python dependencies for the project.  We will also use pipenv as our virtual environment manager even though both pipenv and pyenv both have the capability to spin up venvs.

1 - Install pipenv package and environment manager

```zsh
pip install pipenv
```

2 - install virtual envinronment python import dependencies

```zsh
pipenv install ipykernel numpy pandas tree jupyter -v --clear
```

ipykernel and jupyter are being used for development purposes.  

## Start the Virtual Environment and bind environment to Jupyter kernel

```zsh
pipenv shell
# bind virtual environment / dependency stack to a Jupyter kernel
python -m ipykernel install --user --name=`basename $VIRTUAL_ENV`
```

## Happy Hacking...

You can now either go into Jupyter or launch your favorite IDE / console to code.

```zsh
# run jupyter (AFAIK this is the best method to launch if using UI)
pipenv run jupyter notebook
```