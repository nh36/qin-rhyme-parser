from setuptools import setup


setup(
    name='cldfbench_cjp',
    py_modules=['cldfbench_cjp'],
    include_package_data=True,
    zip_safe=False,
    entry_points={
        'cldfbench.dataset': [
            'cjp=cldfbench_cjp:Dataset',
        ],
        'cldfbench.commands': [
            'cjp=commands',
        ],
    },
    install_requires=[
        'cldfbench',
    ],
    extras_require={
        'test': [
            'pytest-cldf',
        ],
    },
)
