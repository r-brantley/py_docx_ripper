# py_docx_ripper
A Python script that traverses recursively down a set directory starting point to parse Microsoft Word .docx contents

---- we shall clean this up.  we promise.

Recommend using python virtual environments.

brew install pyenv

export PYENV_ROOT=$(pyenv root)
export PATH="$PYENV_ROOT/shims:$PATH"


pyenv local 3.x.x to desired version

brew install pyenv-virtualenv

eval "$(pyenv virtualenv-init -)"




#pip install pipenv first
pip install pipenv 

#instantiate and run venv shell




https://towardsdatascience.com/virtual-environments-for-data-science-running-python-and-jupyter-with-pipenv-c6cb6c44a405

#install import dependencies
pipenv install ipykernel numpy pandas tree jupyter -v --clear
pipenv shell
python -m ipykernel install --user --name=`basename $VIRTUAL_ENV`

pipenv run jupyter notebook


pipenv run jupyter notebook