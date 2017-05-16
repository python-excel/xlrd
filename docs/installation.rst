Installation Instructions
=========================

If you want to experiment with xlrd, the easiest way to
install it is to do the following in a virtualenv::

  pip install xlrd

If your package uses setuptools and you decide to use xlrd,
then you should add it as a requirement by adding an ``install_requires``
parameter in your call to ``setup`` as follows:

.. code-block:: python

    setup(
        # other stuff here
        install_requires=['xlrd'],
        )
