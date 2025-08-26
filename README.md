# acd
ALTHOM Codebase

# INSTALL
pip install git+https://github.com/MihailCosmin/acd


### INSTALL WITH VERSION
python -m pip uninstall -y acd

pip wheel git+https://github.com/MihailCosmin/acd -w dist

##### Install with all dependencies
python -m pip install --force-reinstall --no-index dist/acd-0.0.3.1-py3-none-any.whl

##### Install without dependencies - only update acd
python -m pip install -U --no-deps dist/acd-0.0.3.1-py3-none-any.whl

pip freeze | findstr acd   # should show: acd==0.0.3.1