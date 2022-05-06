# Various linear algebra operations
# - Vector and matrix multiplication
# - Procrustean registration; find rigid rotation to match points
# Usage:
#
# import linear
# reload(linear)
#
# m = linear.get_transform(seq, -1)
# (x,y) = linear.transform(m, seq[i][2], seq[i][3])
# # (x,y) should be approximately (seq[i][0], seq[i][1])

import math

def find_centroids(fixed_and_rotated):
    """
    From a (nonempty) set of quadruplets (x,y,w,z) returns the means of
    each x, y, w, and z in a quadruplet
    """
    n = len(fixed_and_rotated)
    if n == 0:
        return (0, 0, 0, 0)
    xt = 0
    yt = 0
    wt = 0
    zt = 0
    for (x,y,w,z) in fixed_and_rotated:
        xt += x
        yt += y
        wt += w
        zt += z
    return (xt / n, yt / n, wt / n, zt / n)

def find_rotation(fixed_and_rotated):
    """
    From a set of quadruplets (x,y,w,z) where the (x,y)s are the fixed points
    and the (w,z)s are their corresponding rotating points, find the angle that
    best rotates the rotating points close to the fixed points, as a pair of
    its sin and cos.
    """
    n = 0
    d = 0
    for (x,y,w,z) in fixed_and_rotated:
        n += -w*y - z*x
        d += -w*x + z*y
    h = math.hypot(n, d)
    if h < 0.001:
        return (0,1)
    return (n/h, d/h)

def get_transformation(fixed_and_rotated, flip=1):
    """
    Returns the transformation matrix.
    fixed_and_rotated is a sequence of quadruplets (x,y,w,z) where
    (x,y) are the fixed points and (w,z) are the corresponding
    points to be rotated by the matrix returned.
    If flip is -1 the matrix will include a mirroring.
    The return value is [[r0, r1, tx], [r2, r3, ty]] that minimizes
    the errors between the (x,y)s and the
    (r0*w + r1*z + tx, r2*w + r3*z + ty)s
    """
    # RF(x-T) = RFx - RFT
    # F = I or diag(-1, 1), R = (c s / -s c)
    (mpx,mpy,mx,my) = find_centroids(fixed_and_rotated)
    far2 = [(x-mpx,y-mpy,flip*(mx-w),z-my) for (x,y,w,z) in fixed_and_rotated]
    (s,c) = find_rotation(far2)
    return [
        [flip*c, -s, -flip*c*mx + s*my + mpx],
        [flip*s, c, -flip*s*mx - c*my + mpy]
    ]

def table_to_matrix(table):
    if not table:
        return None
    rows = []
    for r in range(table.RowCount):
        column = []
        for c in range(table.ColumnCount):
            column.append(float(table.GetValue(r,c)))
        rows.append(column)
    return rows

def transform(matrix, x, y):
    """
    Returns point (x,y) transformed by matrix
    """
    return (
        x * matrix[0][0] + y * matrix[0][1] + matrix[0][2],
        x * matrix[1][0] + y * matrix[1][1] + matrix[1][2]
    )

def inv_transform(matrix, x, y):
    """
    Returns point (x,y) transformed by the inverse of matrix,
    assuming that matrix is a pure rotation, mirroring and translation
    (that is, an orthogonal matrix)
    """
    xt = x - matrix[0][2]
    yt = y - matrix[1][2]
    return(
        xt * matrix[0][0] + y * matrix[1][0],
        yt * matrix[0][1] + y * matrix[1][1]
    )
