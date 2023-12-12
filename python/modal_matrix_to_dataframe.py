import numpy as np
import pandas as pd

def modal_matrix_to_dataframe(op2, matrix_key):
    """Get a modal matrix indexed by frequency as a dataframe.

    Parameters
    ----------
    op2 : pyNastran.op2.op2.OP2
        OP2 object
    matrix_key : {'EFMFACS', 'MPFACS', 'MEFMASS', 'MEFWTS'}
        key of matrix to return. Assumes SORT1.

    Returns
    -------
    pandas.DataFrame
        Matrix as a DataFrame
    """

    # build dataframes if not already built
    op2.matrices[matrix_key].build_dataframe()
    op2.eigenvectors[1].build_dataframe()

    # create data frame from sparse matrix
    df = pd.DataFrame(
        data=op2.matrices[matrix_key].data.todense().T,
        columns=op2.eigenvectors[1].headers,
        index=op2.eigenvectors[1].data_frame.columns.get_level_values('Freq')
    )

    # add mode numbers as indicies
    df['mode'] = op2.eigenvectors[1].modes
    df = df.reset_index().set_index('mode', drop=True)
    df.columns = [i.capitalize() for i in df.columns]

    return df
