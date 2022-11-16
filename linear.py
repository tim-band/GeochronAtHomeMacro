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
        xt * matrix[0][0] + yt * matrix[1][0],
        xt * matrix[0][1] + yt * matrix[1][1]
    )

# Linear regression for working out focus from stage position assuming
# a flat target
class Regression:
    def __init__(self, table = None):
        """
        Initializes the regression to the data in table (as produced by
        save()) or to no data.
        """
        def value(r, c):
            return float(table.GetValue(r, c))
        if table is not None:
            self.n = value(0,0)
            self.sx = value(0, 1)
            self.sy = value(0, 2)
            self.sz = value(0, 3)
            self.sx2 = value(1, 1)
            self.sxy = value(1, 2)
            self.sxz = value(1, 3)
            self.sy2 = value(2, 2)
            self.syz = value(2, 3)
            self.sz2 = value(3, 3)
        else:
            self.n = 0
            self.sx = 0
            self.sy = 0
            self.sz = 0
            self.sx2 = 0
            self.sy2 = 0
            self.sz2 = 0
            self.sxy = 0
            self.sxz = 0
            self.syz = 0

    def save(self, table):
        """
        Saves the data gathered to a table, deleting
        any data currently stored.
        """
        table.Columns.Clear()
        table.Columns.Add('n')
        table.Columns.Add('x')
        table.Columns.Add('y')
        table.Columns.Add('z')
        table.SetValue(0, 0, self.n)
        table.SetValue(0, 1, self.sx)
        table.SetValue(0, 2, self.sy)
        table.SetValue(0, 3, self.sz)
        table.SetValue(1, 1, self.sx2)
        table.SetValue(1, 2, self.sxy)
        table.SetValue(1, 3, self.sxz)
        table.SetValue(2, 2, self.sy2)
        table.SetValue(2, 3, self.syz)
        table.SetValue(3, 3, self.sz2)

    def count(self):
        """
        Returns the current data point count. Do not use estimate_z()
        if this returns less than 3.
        """
        return self.n

    def add(self, x, y, z):
        """
        Adds z as the correct focus for stage position (x,y)
        """
        self.n += 1
        self.sx += x
        self.sy += y
        self.sz += z
        self.sx2 += x * x
        self.sy2 += y * y
        self.sz2 += z * z
        self.sxy += x * y
        self.sxz += x * z
        self.syz += y * z

    def getXTX(self):
        # X = (1, x0, y0 \ 1, x1, y1 \ 1, x2, y2 \ ...) and Y = (z0 \ z1 \ z1 \ ...)
        # So XTX = (n, sx, sy \ sx, sx2, sxy \ sy, sxy, sy2) and XTY = (sz \ sxz \ syz)
        return [
            [  self.n,  self.sx,  self.sy ],
            [ self.sx, self.sx2, self.sxy ],
            [ self.sy, self.sxy, self.sy2 ]
        ]

    def get_coefficients(self):
        m = self.getXTX()
        mi = invert(m)
        return map(lambda r: r[0]*self.sz + r[1]*self.sxz + r[2]*self.syz, mi)

    def estimate_z(self, x, y):
        """
        Gets the estimated focal position for stage position (x,y)
        """
        b = self.get_coefficients()
        return b[0] + b[1]*x + b[2]*y

def eliminate(m, r, c, using_r):
    d = m[using_r][c]
    k = m[r][c]/d
    m[r] = map(lambda v,u: v - k * u, m[r], m[using_r])

def finalize(m, r):
    k = m[r][r]
    return [m[r][3]/k, m[r][4]/k, m[r][5]/k]

def invert(m):
    mi = [m[0] + [1,0,0], m[1] + [0,1,0], m[2] + [0,0,1]]
    eliminate(mi, 1, 0, 0)
    eliminate(mi, 2, 0, 0)
    if abs(m[1][1]) < abs(m[0][1]):
        eliminate(mi, 2, 1, 0)
    else:
        eliminate(mi, 2, 1, 1)
    eliminate(mi, 0, 1, 1)
    eliminate(mi, 0, 2, 2)
    eliminate(mi, 1, 2, 2)
    return [finalize(mi, 0), finalize(mi, 1), finalize(mi, 2)]
