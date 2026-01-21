#!/usr/bin/env python
# coding: utf-8

# In[1]:


pip install etabs-api


# In[2]:


import comtypes.client
import pandas as pd
import numpy as np


# In[3]:


def connect_to_etabs():
    #create API helper object
    helper = comtypes.client.CreateObject('ETABSv1.Helper');
    helper = helper.QueryInterface(comtypes.gen.ETABSv1.cHelper);
    #attach to a running instance of ETABS
    try:
        #get the active ETABS object
        myETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject");
        print("Connected to ETABS model");
    except (OSError, comtypes.COMError):
        print("No running instance of the program found or failed to attach.");
        sys.exit(-1);
    #create SapModel object
    SapModel = myETABSObject.SapModel;
    return SapModel,myETABSObject,helper;


# In[4]:


SapModel, myETABSObject, helper = connect_to_etabs()


# In[5]:


import win32com.client
import comtypes.client as ct


# In[6]:


# Connect to the ETABS application
app = ct.GetActiveObject("CSI.ETABS.API.ETABSObject")
model = app.SapModel


# In[7]:


table = model.DatabaseTables.GetTableForDisplayArray("Point Object Connectivity",GroupName="")
#print(table)
cols = table[2]

noOfRows = table[3]

vals = np.array_split(table[4],noOfRows)
XYZ1 = pd.DataFrame(vals)

XYZ1.columns = cols
#display(XYZ1)
#print(XYZ1)

columns_to_remove = ["IsAuto", "Story", "PointBay", "IsSpecial", "GUID"]
#Drop the specified columns from the DataFrame
XYZ1 = XYZ1.drop(columns=columns_to_remove)

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
# Display the modified table without the specified columns
print(XYZ1)


# In[8]:


table = model.DatabaseTables.GetTableForDisplayArray("Beam Object Connectivity",GroupName="")
#print(table)
cols = table[2]

noOfRows = table[3]

vals = np.array_split(table[4],noOfRows)
beam = pd.DataFrame(vals)

beam.columns = cols
#print(beam)

columns_to_remove = ["UniqueName", "Story", "Length", "GUID"]
#Drop the specified columns from the DataFrame
beam = beam.drop(columns=columns_to_remove)


# Display the modified table without the specified columns
print(beam)


# In[9]:


table = model.DatabaseTables.GetTableForDisplayArray("Column Object Connectivity",GroupName="")
#print(table)
cols = table[2]

noOfRows = table[3]

vals = np.array_split(table[4],noOfRows)
column = pd.DataFrame(vals)

column.columns = cols
#print(column)

columns_to_remove = ["UniqueName", "Length", "GUID"]
#Drop the specified columns from the DataFrame
column = column.drop(columns=columns_to_remove)


# Display the modified table without the specified columns
print(column)


# In[10]:


table = model.DatabaseTables.GetTableForDisplayArray("Floor Object Connectivity",GroupName="")
#print(table)
cols = table[2]

noOfRows = table[3]

vals = np.array_split(table[4],noOfRows)
floor = pd.DataFrame(vals)

floor.columns = cols
#print(floor)

columns_to_remove = ["UniqueName", "Story", "Perimeter", "GUID"]
#Drop the specified columns from the DataFrame
floor = floor.drop(columns=columns_to_remove)


# Display the modified table without the specified columns
print(floor)


# In[11]:


import pandas as pd

D = pd.DataFrame(floor)

# Function to extract unique points from the 'UniquePts' column
def extract_unique_pts(row, index):
    try:
        return int(row['UniquePts'].split('; ')[index])
    except IndexError:
        return None

# Extract individual points into separate columns
for i in range(1, 5):  # Assuming there are at most 4 points (adjust as needed)
    column_name = f'UniquePt{i}'
    D[column_name] = D.apply(lambda row: extract_unique_pts(row, i-1), axis=1)

# Print the updated DataFrame
print(D)


# In[12]:


get_ipython().run_cell_magic('capture', '', "import pandas as pd\nimport matplotlib.pyplot as plt\nfrom mpl_toolkits.mplot3d import Axes3D\n\n# Convert UniqueName to strings\nA = pd.DataFrame(XYZ1)\n\nA['UniqueName'] = A['UniqueName'].astype(str)\n\nA['X'] = pd.to_numeric(A['X'])\nA['Y'] = pd.to_numeric(A['Y'])\nA['Z'] = pd.to_numeric(A['Z'])\n\nA['X'] = A['X']/1000\nA['Y'] = A['Y']/1000\nA['Z'] = A['Z']/1000\n\n# Create a 3D plot\nfig = plt.figure()\nax = fig.add_subplot(111, projection='3d')\n\n# Plot points with labels\nfor i, row in A.iterrows():\n    ax.scatter(row['X'], row['Y'], row['Z'], label=row['UniqueName'])\n\n# Customize plot if needed\nax.set_xlabel('X-axis')\nax.set_ylabel('Y-axis')\nax.set_zlabel('Z-axis')\nax.legend()\n\n# Show the plot\n\ndef set_axes_equal(ax):\n    x_limits = ax.get_xlim3d()\n    y_limits = ax.get_ylim3d()\n    z_limits = ax.get_zlim3d()\n\n    x_range = abs(x_limits[1] - x_limits[0])\n    y_range = abs(y_limits[1] - y_limits[0])\n    z_range = abs(z_limits[1] - z_limits[0])\n\n    max_range = max(x_range, y_range, z_range)\n\n    mid_x = (x_limits[0] + x_limits[1]) / 2\n    mid_y = (y_limits[0] + y_limits[1]) / 2\n    mid_z = (z_limits[0] + z_limits[1]) / 2\n\n    ax.set_xlim3d([mid_x - max_range/2, mid_x + max_range/2])\n    ax.set_ylim3d([mid_y - max_range/2, mid_y + max_range/2])\n    ax.set_zlim3d([mid_z - max_range/2, mid_z + max_range/2])\nset_axes_equal(ax)\nplt.show()\nplt.close()\n")


# In[13]:


get_ipython().run_cell_magic('capture', '', "import pandas as pd\nimport matplotlib.pyplot as plt\nfrom mpl_toolkits.mplot3d import Axes3D\n\n# Assuming DataFrames A and B are already defined\n\nB = pd.DataFrame(beam)\n\n# Merge DataFrames A and B on UniqueName to get coordinates for UniquePtI and UniquePtJ\nmerged_df = pd.merge(B, A, left_on='UniquePtI', right_on='UniqueName', how='left', suffixes=('_B', '_A'))\nmerged_df = pd.merge(merged_df, A, left_on='UniquePtJ', right_on='UniqueName', how='left', suffixes=('_B', '_A'))\n\n\n# Create a 3D plot\nfig = plt.figure()\nax = fig.add_subplot(111, projection='3d')\n\n# Plot beams connecting points\nfor i, row in merged_df.iterrows():\n    x = [row['X_A'], row['X_B']]\n    y = [row['Y_A'], row['Y_B']]\n    z = [row['Z_A'], row['Z_B']]\n    ax.plot(x, y, z, label=row['BeamBay'])\n\n# Plot points with labels\nfor i, row in A.iterrows():\n    ax.scatter(row['X'], row['Y'], row['Z'], label=row['UniqueName'])\n\n# Customize plot if needed\nax.set_xlabel('X-axis')\nax.set_ylabel('Y-axis')\nax.set_zlabel('Z-axis')\nax.legend()\n\n# Show the plot\ndef set_axes_equal(ax):\n    x_limits = ax.get_xlim3d()\n    y_limits = ax.get_ylim3d()\n    z_limits = ax.get_zlim3d()\n\n    x_range = abs(x_limits[1] - x_limits[0])\n    y_range = abs(y_limits[1] - y_limits[0])\n    z_range = abs(z_limits[1] - z_limits[0])\n\n    max_range = max(x_range, y_range, z_range)\n\n    mid_x = (x_limits[0] + x_limits[1]) / 2\n    mid_y = (y_limits[0] + y_limits[1]) / 2\n    mid_z = (z_limits[0] + z_limits[1]) / 2\n\n    ax.set_xlim3d([mid_x - max_range/2, mid_x + max_range/2])\n    ax.set_ylim3d([mid_y - max_range/2, mid_y + max_range/2])\n    ax.set_zlim3d([mid_z - max_range/2, mid_z + max_range/2])\nset_axes_equal(ax)\nplt.show()\n")


# In[14]:


get_ipython().run_cell_magic('capture', '', "import pandas as pd\nimport matplotlib.pyplot as plt\nfrom mpl_toolkits.mplot3d import Axes3D\n\n# Assuming DataFrames A and C are already defined\n\nC = pd.DataFrame(column)\n\n# Assuming DataFrames A, B, and C are already defined\n\n# Merge DataFrames A, B, and C on UniqueName to get coordinates for UniquePtI and UniquePtJ\nmerged_df_bc = pd.merge(B, A, left_on='UniquePtI', right_on='UniqueName', how='left', suffixes=('_B', '_A'))\nmerged_df_bc = pd.merge(merged_df_bc, A, left_on='UniquePtJ', right_on='UniqueName', how='left', suffixes=('_B', '_A'))\n\nmerged_df_c = pd.merge(C, A, left_on='UniquePtI', right_on='UniqueName', how='left', suffixes=('_C', '_A'))\nmerged_df_c = pd.merge(merged_df_c, A, left_on='UniquePtJ', right_on='UniqueName', how='left', suffixes=('_C', '_A'))\n\nprint(merged_df_c)\n\n# Create a 3D plot\n#fig = plt.figure()\nfig = plt.figure(figsize=(20, 20))\nax = fig.add_subplot(111, projection='3d')\n\n# Plot beams\nfor i, row in merged_df_bc.iterrows():\n    x = [row['X_A'], row['X_B']]\n    y = [row['Y_A'], row['Y_B']]\n    z = [row['Z_A'], row['Z_B']]\n    ax.plot(x, y, z, label=row['BeamBay'])\n\n# Plot columns\nfor i, row in merged_df_c.iterrows():\n    x = [row['X_A'], row['X_C']]\n    y = [row['Y_A'], row['Y_C']]\n    z = [row['Z_A'], row['Z_C']]\n    ax.plot(x, y, z, label=row['ColumnBay'])\n\n# Plot points with labels\nfor i, row in A.iterrows():\n    ax.scatter(row['X'], row['Y'], row['Z'], label=row['UniqueName'])\n\n# Customize plot if needed\nax.set_xlabel('X-axis')\nax.set_ylabel('Y-axis')\nax.set_zlabel('Z-axis')\nax.legend()\n\n# Show the plot\ndef set_axes_equal(ax):\n    x_limits = ax.get_xlim3d()\n    y_limits = ax.get_ylim3d()\n    z_limits = ax.get_zlim3d()\n\n    x_range = abs(x_limits[1] - x_limits[0])\n    y_range = abs(y_limits[1] - y_limits[0])\n    z_range = abs(z_limits[1] - z_limits[0])\n\n    max_range = max(x_range, y_range, z_range)\n\n    mid_x = (x_limits[0] + x_limits[1]) / 2\n    mid_y = (y_limits[0] + y_limits[1]) / 2\n    mid_z = (z_limits[0] + z_limits[1]) / 2\n\n    ax.set_xlim3d([mid_x - max_range/2, mid_x + max_range/2])\n    ax.set_ylim3d([mid_y - max_range/2, mid_y + max_range/2])\n    ax.set_zlim3d([mid_z - max_range/2, mid_z + max_range/2])\nset_axes_equal(ax)\nplt.show()\n")


# In[15]:


get_ipython().run_cell_magic('capture', '', 'import pandas as pd\nimport matplotlib.pyplot as plt\nfrom mpl_toolkits.mplot3d import Axes3D  # noqa: F401 (needed for 3D)\nfrom mpl_toolkits.mplot3d.art3d import Poly3DCollection\nimport numpy as np\n\n# --- Assumptions ---\n# A has: [\'UniqueName\',\'X\',\'Y\',\'Z\']  -> master node list\n# B has: [\'UniquePtI\',\'UniquePtJ\',\'BeamBay\']\n# C has: [\'UniquePtI\',\'UniquePtJ\',\'ColumnBay\']\n# D has: [\'UniquePts\',\'FloorBay\'] where UniquePts = "12; 15; 18; 22" (ordered boundary)\n\n# ---------- Helpers ----------\ndef _to_key(x):\n    """Normalize point id to string key (handles int/str and stray spaces)."""\n    if pd.isna(x):\n        return None\n    return str(x).strip()\n\ndef parse_unique_pts(s):\n    """Return ordered list of normalized point ids from \'a; b; c; ...\'."""\n    if pd.isna(s):\n        return []\n    parts = [p.strip() for p in str(s).split(\';\')]\n    return [_to_key(p) for p in parts if p.strip() != ""]\n\ndef build_coord_lookup(A):\n    """Build {UniqueName(str): (x,y,z)} lookup from A."""\n    # Normalize key as string to be consistent\n    A_loc = A.copy()\n    A_loc[\'__key__\'] = A_loc[\'UniqueName\'].apply(_to_key)\n    return {\n        k: (row[\'X\'], row[\'Y\'], row[\'Z\'])\n        for k, row in A_loc.set_index(\'__key__\')[[\'X\',\'Y\',\'Z\']].iterrows()\n    }\n\ndef coords_for_ids(ids, coord_lu):\n    """Get list of (x,y,z) for given id list; skip missing ids."""\n    out = []\n    for pid in ids:\n        if pid in coord_lu:\n            out.append(coord_lu[pid])\n    return out\n\ndef set_axes_equal(ax):\n    """Equal aspect for 3D axes."""\n    x_limits = ax.get_xlim3d()\n    y_limits = ax.get_ylim3d()\n    z_limits = ax.get_zlim3d()\n    x_range = abs(x_limits[1] - x_limits[0])\n    y_range = abs(y_limits[1] - y_limits[0])\n    z_range = abs(z_limits[1] - z_limits[0])\n    plot_radius = 0.5 * max([x_range, y_range, z_range])\n\n    x_middle = np.mean(x_limits)\n    y_middle = np.mean(y_limits)\n    z_middle = np.mean(z_limits)\n    ax.set_xlim3d([x_middle - plot_radius, x_middle + plot_radius])\n    ax.set_ylim3d([y_middle - plot_radius, y_middle + plot_radius])\n    ax.set_zlim3d([z_middle - plot_radius, z_middle + plot_radius])\n\n# ---------- Prep ----------\n# Build coordinate lookup from A\ncoord_lu = build_coord_lookup(A)\n\n# Normalize B/C point ids to strings for lookup\nfor df in (B, C):\n    df[\'__I__\'] = df[\'UniquePtI\'].apply(_to_key)\n    df[\'__J__\'] = df[\'UniquePtJ\'].apply(_to_key)\n\n# Parse polygon vertex ids for each slab row\nD = D.copy()\nD[\'__ids__\'] = D[\'UniquePts\'].apply(parse_unique_pts)\n\n# ---------- Plot ----------\nfig = plt.figure(figsize=(18, 14))\nax = fig.add_subplot(111, projection=\'3d\')\n\n# Beams (lines)\nfor _, row in B.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        ax.plot([x1, x2], [y1, y2], [z1, z2], label=None)\n\n# Columns (lines)\nfor _, row in C.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        ax.plot([x1, x2], [y1, y2], [z1, z2], label=None)\n\n# Slabs (any polygon)\nfaces = []\nedge_color = \'k\'   # you can change if you like\nface_alpha = 0.45  # translucency\n\nfor _, row in D.iterrows():\n    ids = row[\'__ids__\']\n    verts = coords_for_ids(ids, coord_lu)\n    if len(verts) < 3:\n        continue  # need at least a triangle\n\n    # Draw polygon edges (closed loop)\n    xs, ys, zs = zip(*(verts + [verts[0]]))\n    ax.plot(xs, ys, zs, color=edge_color, linewidth=1.5)\n\n    # Collect for filled faces\n    faces.append(verts)\n\n# Add filled faces in one go (faster)\nif faces:\n    poly = Poly3DCollection(faces, alpha=face_alpha)\n    # NOTE: not setting a color -> matplotlib will auto-pick; specify if you want\n    poly.set_edgecolor(edge_color)\n    ax.add_collection3d(poly)\n\n# Cosmetics\nax.set_xlabel(\'X\')\nax.set_ylabel(\'Y\')\nax.set_zlabel(\'Z\')\nax.set_title(\'3D Frame with Slabs (Arbitrary Polygon Support)\')\n\n# Make aspect equal-ish\nset_axes_equal(ax)\n\n\nplt.tight_layout()\n#plt.show()\n')


# In[16]:


import pandas as pd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401 (needed for 3D)
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
import numpy as np

# --- Assumptions ---
# A has: ['UniqueName','X','Y','Z']  -> master node list
# B has: ['UniquePtI','UniquePtJ','BeamBay']
# C has: ['UniquePtI','UniquePtJ','ColumnBay']
# D has: ['UniquePts','FloorBay'] where UniquePts = "12; 15; 18; 22" (ordered boundary)

# ---------- Helpers ----------
def _to_key(x):
    """Normalize point id to string key (handles int/str and stray spaces)."""
    if pd.isna(x):
        return None
    return str(x).strip()

def parse_unique_pts(s):
    """Return ordered list of normalized point ids from 'a; b; c; ...'."""
    if pd.isna(s):
        return []
    parts = [p.strip() for p in str(s).split(';')]
    return [_to_key(p) for p in parts if p.strip() != ""]

def build_coord_lookup(A):
    """Build {UniqueName(str): (x,y,z)} lookup from A."""
    # Normalize key as string to be consistent
    A_loc = A.copy()
    A_loc['__key__'] = A_loc['UniqueName'].apply(_to_key)
    return {
        k: (row['X'], row['Y'], row['Z'])
        for k, row in A_loc.set_index('__key__')[['X','Y','Z']].iterrows()
    }

def coords_for_ids(ids, coord_lu):
    """Get list of (x,y,z) for given id list; skip missing ids."""
    out = []
    for pid in ids:
        if pid in coord_lu:
            out.append(coord_lu[pid])
    return out

def set_axes_equal(ax):
    """Equal aspect for 3D axes."""
    x_limits = ax.get_xlim3d()
    y_limits = ax.get_ylim3d()
    z_limits = ax.get_zlim3d()
    x_range = abs(x_limits[1] - x_limits[0])
    y_range = abs(y_limits[1] - y_limits[0])
    z_range = abs(z_limits[1] - z_limits[0])
    plot_radius = 0.5 * max([x_range, y_range, z_range])

    x_middle = np.mean(x_limits)
    y_middle = np.mean(y_limits)
    z_middle = np.mean(z_limits)
    ax.set_xlim3d([x_middle - plot_radius, x_middle + plot_radius])
    ax.set_ylim3d([y_middle - plot_radius, y_middle + plot_radius])
    ax.set_zlim3d([z_middle - plot_radius, z_middle + plot_radius])

# ---------- Prep ----------
# Build coordinate lookup from A
coord_lu = build_coord_lookup(A)

# Normalize B/C point ids to strings for lookup
for df in (B, C):
    df['__I__'] = df['UniquePtI'].apply(_to_key)
    df['__J__'] = df['UniquePtJ'].apply(_to_key)

# Parse polygon vertex ids for each slab row
D = D.copy()
D['__ids__'] = D['UniquePts'].apply(parse_unique_pts)

# ---------- Plot ----------
fig = plt.figure(figsize=(18, 14))
ax = fig.add_subplot(111, projection='3d')

# Beams (lines)
for _, row in B.iterrows():
    pI, pJ = row['__I__'], row['__J__']
    if pI in coord_lu and pJ in coord_lu:
        (x1, y1, z1) = coord_lu[pI]
        (x2, y2, z2) = coord_lu[pJ]
        ax.plot([x1, x2], [y1, y2], [z1, z2], label=None)

# Columns (lines)
for _, row in C.iterrows():
    pI, pJ = row['__I__'], row['__J__']
    if pI in coord_lu and pJ in coord_lu:
        (x1, y1, z1) = coord_lu[pI]
        (x2, y2, z2) = coord_lu[pJ]
        ax.plot([x1, x2], [y1, y2], [z1, z2], label=None)

# Slabs (any polygon)
faces = []
edge_color = 'k'   # you can change if you like
face_alpha = 0.45  # translucency

for _, row in D.iterrows():
    ids = row['__ids__']
    verts = coords_for_ids(ids, coord_lu)
    if len(verts) < 3:
        continue  # need at least a triangle

    # Draw polygon edges (closed loop)
    xs, ys, zs = zip(*(verts + [verts[0]]))
    ax.plot(xs, ys, zs, color=edge_color, linewidth=1.5)

    # Collect for filled faces
    faces.append(verts)

# Add filled faces in one go (faster)
if faces:
    poly = Poly3DCollection(faces, alpha=face_alpha)
    # NOTE: not setting a color -> matplotlib will auto-pick; specify if you want
    poly.set_edgecolor(edge_color)
    ax.add_collection3d(poly)

# Cosmetics
ax.set_xlabel('X')
ax.set_ylabel('Y')
ax.set_zlabel('Z')
ax.set_title('3D Frame with Slabs (Arbitrary Polygon Support)')

# Make aspect equal-ish
set_axes_equal(ax)

plt.tight_layout()
#plt.show()



# draw stars--- detect coordinate column names robustly ---
cols = merged_df_c.columns.tolist()

# possible coordinate patterns to look for
candidates_X = [c for c in cols if c.startswith('X_') or c.startswith('X')]
candidates_Y = [c for c in cols if c.startswith('Y_') or c.startswith('Y')]
candidates_Z = [c for c in cols if c.startswith('Z_') or c.startswith('Z')]

# prefer explicit suffixes if present
def find_pair(suffix1, suffix2):
    x1, y1, z1 = f'X{suffix1}', f'Y{suffix1}', f'Z{suffix1}'
    x2, y2, z2 = f'X{suffix2}', f'Y{suffix2}', f'Z{suffix2}'
    if all(x in cols for x in (x1,y1,z1,x2,y2,z2)):
        return (x1,y1,z1), (x2,y2,z2)
    return None

# common suffix combos to try (based on your printout)
pairs_to_try = [('_A','_C'), ('_C','_A'), ('_bottom','_top'), ('_top','_bottom'), ('_A','_B'), ('','_top')]

bottom_cols = top_cols = None
for s1, s2 in pairs_to_try:
    found = find_pair(s1, s2)
    if found:
        bottom_cols, top_cols = found
        break

# fallback: try to pick any two different X_/Z_ columns
if bottom_cols is None:
    x_cols = [c for c in cols if c.startswith('X_')]
    z_cols = [c for c in cols if c.startswith('Z_')]
    if len(x_cols) >= 2 and len(z_cols) >= 2:
        # choose first two distinct suffixes
        # map by suffix extracted after '_'
        suf_list = list({c.split('_',1)[1] for c in x_cols})
        if len(suf_list) >= 2:
            s1, s2 = suf_list[0], suf_list[1]
            bottom_cols = (f'X_{s1}', f'Y_{s1}', f'Z_{s1}')
            top_cols    = (f'X_{s2}', f'Y_{s2}', f'Z_{s2}')

# if still None, raise helpful error
if bottom_cols is None or top_cols is None:
    raise KeyError(f"Couldn't auto-detect bottom/top coordinate columns. Available cols: {cols}")

xb_col, yb_col, zb_col = bottom_cols
xt_col, yt_col, zt_col = top_cols

# check which is actually the top (higher Z on average)
mean_zb = merged_df_c[zb_col].mean()
mean_zt = merged_df_c[zt_col].mean()

if mean_zt < mean_zb:
    # swap if our naming guess was reversed
    xb_col, xt_col = xt_col, xb_col
    yb_col, yt_col = yt_col, yb_col
    zb_col, zt_col = zt_col, zb_col

print(f"Using bottom coords: {xb_col}, {yb_col}, {zb_col}")
print(f"Using top    coords: {xt_col}, {yt_col}, {zt_col}")
print(f"mean Z bottom = {merged_df_c[zb_col].mean():.3f}, mean Z top = {merged_df_c[zt_col].mean():.3f}")

# --- Plot columns and star at top ---
for i, row in merged_df_c.iterrows():
    x = [row[xb_col], row[xt_col]]
    y = [row[yb_col], row[yt_col]]
    z = [row[zb_col], row[zt_col]]
    ax.plot(x, y, z, color='black', label='Column' if i == 0 else '')

    # optionally offset the star slightly above top to sit above slab
    offset = 0.05  # change units as appropriate (meters or same units as Z)
    ax.scatter(row[xt_col], row[yt_col], row[zt_col] + offset,
               color='red', marker='*', s=200, edgecolor='k', linewidth=0.8)
    
    
      # --- Dark-blue neighbor lines on every floor (NO FILTERING) ---
import numpy as np

# 1) Group columns by floor using top-Z
floor_decimals = 3
merged_df_c['_floor_id_'] = np.round(merged_df_c[zt_col].to_numpy(), floor_decimals)

# 2) Settings
k_nearest    = None      # None => connect to ALL other same-floor columns
max_radius   = None      # e.g., 12.0 to limit distance; keep None to ignore

# 3) Avoid duplicate segments across pairs
_drawn_pairs = set()  # store frozenset((idx_i, idx_j))

# 4) For each floor, connect columns
for fid, df_floor in merged_df_c.groupby('_floor_id_'):
    if df_floor.shape[0] < 2:
        continue

    idxs = df_floor.index.to_list()
    XY   = df_floor[[xt_col, yt_col]].to_numpy()
    Z    = df_floor[zt_col].to_numpy()

    for i, idx_i in enumerate(idxs):
        x0, y0, z0 = XY[i,0], XY[i,1], Z[i]

        # distances to all others on the same floor
        d2 = np.sum((XY - XY[i])**2, axis=1)
        d  = np.sqrt(d2)

        # sort by distance, skip self
        order = np.argsort(d)
        order = [j for j in order if j != i]

        # optional radius cap
        if max_radius is not None:
            order = [j for j in order if d[j] <= max_radius]

        # take k nearest or all
        if k_nearest is not None:
            order = order[:k_nearest]
        if not order:
            continue

        neighbors_df = df_floor.iloc[order].copy()
        neighbors_df['_distXY_'] = d[order]

        # draw every undirected pair exactly once (no hidden/collinear filtering)
        for j in order:
            idx_j = idxs[j]
            if frozenset((idx_i, idx_j)) in _drawn_pairs:
                continue
            _drawn_pairs.add(frozenset((idx_i, idx_j)))

            x1, y1, z1 = XY[j,0], XY[j,1], Z[j]

            # dark-blue connection
            ax.plot([x0, x1], [y0, y1], [z0, z1],
                    color='#003366', linewidth=4.0, linestyle=':')
            
            
              # --- midpoint (black)  |  FIXED: ym uses y1, not x1 ---
            xm = (x0 + x1) / 2.0
            ym = (y0 + y1) / 2.0   # <-- bug fix here
            zm = (z0 + z1) / 2.0
            ax.scatter(xm, ym, zm, color='black', s=50, zorder=12)
            
            
             # --- XY-plane perpendicular through midpoint (dark orange) ---
            dx = x1 - x0
            dy = y1 - y0
            seg_len = (dx*dx + dy*dy) ** 0.5
            if seg_len > 0:
                # unit perpendicular in XY
                px = -dy / seg_len
                py =  dx / seg_len

                # length = 25% of the blue segment (tweak as needed)
                half_len = 0.25 * seg_len
                xL, yL = xm - px * half_len, ym - py * half_len
                xR, yR = xm + px * half_len, ym + py * half_len

                # keep Z constant on the floor
                ax.plot([xL, xR], [yL, yR], [zm,  zm],
                        color='#FF8C00', linewidth=3.0, zorder=13)
                
                
                # ========== SHADE VORONOI-LIKE REGIONS PER COLUMN, CLIPPED TO ACTUAL SLAB POLYGONS (INCLUDES OPENINGS) ==========
import numpy as np
from collections import defaultdict
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from matplotlib import colormaps

# Use shapely for robust polygon / area / intersection handling
try:
    from shapely.geometry import Polygon, Point, box, MultiPolygon
    from shapely.ops import unary_union
except Exception as e:
    raise ImportError("This script requires shapely. Install with `pip install shapely` and re-run.") from e

# --- helpers ---
def convex_hull(points_xy):
    """Monotone chain convex hull; returns hull [(x,y),...] CCW."""
    pts = np.unique(np.asarray(points_xy, dtype=float), axis=0)
    if len(pts) <= 2:
        return pts.tolist()
    pts = pts[np.lexsort((pts[:,1], pts[:,0]))]
    def cross(o, a, b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    lower = []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(tuple(p))
    upper = []
    for p in pts[::-1]:
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(tuple(p))
    return lower[:-1] + upper[:-1]

def clip_halfplane(poly, m, n):
    """Sutherland–Hodgman clip: keep side (p - m)·n <= 0."""
    if not poly:
        return []
    out = []
    def inside(p):
        return ((p[0]-m[0])*n[0] + (p[1]-m[1])*n[1]) <= 1e-9
    def intersect(p1, p2):
        dp = (p2[0]-p1[0], p2[1]-p1[1])
        num = (m[0]-p1[0])*n[0] + (m[1]-p1[1])*n[1]
        den = dp[0]*n[0] + dp[1]*n[1]
        if abs(den) < 1e-12:
            return p2
        t = num/den
        return (p1[0]+t*dp[0], p1[1]+t*dp[1])
    S = poly[-1]
    Sin = inside(S)
    for E in poly:
        Ein = inside(E)
        if Ein:
            if not Sin:
                out.append(intersect(S, E))
            out.append(E)
        else:
            if Sin:
                out.append(intersect(S, E))
        S, Sin = E, Ein
    return out

def polygon_area_xy(poly):
    """Shoelace area for a 2D polygon [(x,y),...]."""
    if not poly or len(poly) < 3:
        return 0.0
    x = np.array([p[0] for p in poly], dtype=float)
    y = np.array([p[1] for p in poly], dtype=float)
    return 0.5 * abs(np.dot(x, np.roll(y, -1)) - np.dot(y, np.roll(x, -1)))

# --- 0) sanity: required names exist
# Required inputs: D (slab rows with '__ids__'), coord_lu (id -> (x,y,z)), merged_df_c, xt_col, yt_col, zt_col, ax
# If _floor_id_ not present, create from zt_col using floor_decimals
floor_decimals = 3
if '_floor_id_' not in merged_df_c.columns:
    merged_df_c['_floor_id_'] = np.round(merged_df_c[zt_col].to_numpy(), floor_decimals)

# --- 1) Build accurate per-floor slab geometry (using shapely) including cutouts/openings ---
# Assumption: each row in D corresponds to a slab polygon (outer boundary). If your slab rows explicitly encode holes/openings
# as separate rows, union will take care of them. If openings are encoded differently you'll need to adapt how rows are grouped.
floor_polygons = {}   # fid -> shapely Polygon or MultiPolygon
floor_points = defaultdict(list)
floor_zlevel = {}

for _, row in D.iterrows():
    ids = row.get('__ids__', [])
    verts = [coord_lu.get(i) for i in ids if i in coord_lu]
    verts = [v for v in verts if v is not None]
    if len(verts) < 3:
        continue
    xs, ys, zs = zip(*verts)
    fid = float(np.round(np.median(zs), floor_decimals))
    floor_points[fid].extend((float(x), float(y)) for x, y in zip(xs, ys))
    floor_zlevel[fid] = float(np.median(zs))
    try:
        # create polygon from the vertex loop (assumed ordered). If not ordered, consider convex_hull fallback.
        poly = Polygon([(float(x), float(y)) for x, y in zip(xs, ys)])
        if not poly.is_valid:
            # try to fix small issues
            poly = poly.buffer(0)
        if poly.is_valid and not poly.is_empty:
            if fid not in floor_polygons:
                floor_polygons[fid] = []
            floor_polygons[fid].append(poly)
    except Exception:
        # fallback: skip invalid polygon
        continue

# union all polygons per floor to get the real slab region (handles multiple slab pieces and holes)
for fid, polys in list(floor_polygons.items()):
    if not polys:
        continue
    u = unary_union(polys)
    # ensure we store either Polygon or MultiPolygon
    floor_polygons[fid] = u

# If any floor has no explicit slabs in D, fallback to convex hull of points for that floor (preserves earlier behaviour)
for fid, pts in floor_points.items():
    if fid not in floor_polygons or floor_polygons[fid] is None or floor_polygons[fid].is_empty:
        hull = convex_hull(pts)
        if len(hull) >= 3:
            floor_polygons[fid] = Polygon(hull)

# --- 2) Shade Voronoi-like cells per floor, but clip each cell to the true slab polygon (so openings are respected) ---
cmap = colormaps.get_cmap('tab20')

for fid, df_floor in merged_df_c.groupby('_floor_id_'):
    if df_floor.shape[0] == 0:
        continue

    XY = df_floor[[xt_col, yt_col]].to_numpy(dtype=float)
    Z  = df_floor[zt_col].to_numpy(dtype=float)
    z_floor = float(np.median(Z)) if fid not in floor_zlevel else floor_zlevel[fid]

    # robust slab polygon for this floor (may be Polygon or MultiPolygon). If missing, create padded bbox from XY.
    slab_poly = floor_polygons.get(fid, None)
    if slab_poly is None or slab_poly.is_empty:
    # NO SLAB → tributary area = zero
        merged_df_c.loc[df_floor.index, 'CellArea'] = 0.0
        continue


    # We'll still use half-plane clipping to compute candidate Voronoi cell, then intersect with slab_poly
    ncols = len(df_floor)
    for i in range(ncols):
        si = XY[i]
        # start with a large bbox around slab to ensure clipping numerics are stable
        minx, miny, maxx, maxy = slab_poly.bounds
        padx = 0.2*(maxx-minx if maxx>minx else 1.0)
        pady = 0.2*(maxy-miny if maxy>miny else 1.0)
        bounds_poly = [(minx-padx,miny-pady),(maxx+padx,miny-pady),
                       (maxx+padx,maxy+pady),(minx-padx,maxy+pady)]
        cell = bounds_poly[:]
        for j in range(ncols):
            if j == i:
                continue
            sj = XY[j]
            m = 0.5*(si + sj)
            n = (sj - si)
            cell = clip_halfplane(cell, m, n)
            if not cell:
                break
        if not cell:
            continue

        # intersect the clipped half-plane polygon with the true slab polygon to remove areas inside cutouts/openings
        try:
            cell_poly = Polygon(cell)
            if not cell_poly.is_valid:
                cell_poly = cell_poly.buffer(0)
            clipped = cell_poly.intersection(slab_poly)
        except Exception:
            # if shapely fails for some reason, skip plotting this site
            continue

        if clipped.is_empty:
            continue

        # Plot potentially multiple polygon parts (MultiPolygon) as separate faces at z_floor
        if isinstance(clipped, (MultiPolygon, )):
            parts = list(clipped.geoms)
        else:
            parts = [clipped]

        color = cmap((i % 20)/20.0)
        for part in parts:
            if not part.exterior:
                continue
            coords2d = list(part.exterior.coords)
            verts3d = [(x, y, z_floor) for (x, y) in coords2d]
            poly = Poly3DCollection([verts3d], facecolors=color, edgecolors='k',
                                    linewidths=0.6, alpha=0.35)
            ax.add_collection3d(poly)

# --- 3) Compute/attach cell area for each column (XY area) using intersection with actual slab polygons (no scaling) ---
merged_df_c['CellArea'] = np.nan

for fid, df_floor in merged_df_c.groupby('_floor_id_'):
    XY = df_floor[[xt_col, yt_col]].to_numpy(dtype=float)

    slab_poly = floor_polygons.get(fid, None)
    if slab_poly is None or slab_poly.is_empty:
    # NO SLAB → tributary area = zero
        merged_df_c.loc[df_floor.index, 'CellArea'] = 0.0
        continue


    idxs = df_floor.index.to_list()
    ncols = len(df_floor)

    for i in range(ncols):
        si = XY[i]
        minx, miny, maxx, maxy = slab_poly.bounds
        padx = 0.2*(maxx-minx if maxx>minx else 1.0)
        pady = 0.2*(maxy-miny if maxy>miny else 1.0)
        bounds_poly = [(minx-padx,miny-pady),(maxx+padx,miny-pady),
                       (maxx+padx,maxy+pady),(minx-padx,maxy+pady)]
        cell = bounds_poly[:]
        for j in range(ncols):
            if j == i:
                continue
            sj = XY[j]
            m = 0.5*(si + sj)
            n = (sj - si)
            cell = clip_halfplane(cell, m, n)
            if not cell:
                break
        if not cell:
            area = 0.0
        else:
            try:
                cell_poly = Polygon(cell)
                if not cell_poly.is_valid:
                    cell_poly = cell_poly.buffer(0)
                clipped = cell_poly.intersection(slab_poly)
                # shapely area is exact XY area
                area = float(clipped.area) if not clipped.is_empty else 0.0
            except Exception:
                area = 0.0
        merged_df_c.loc[idxs[i], 'CellArea'] = area

# --- 4) NO SCALING: TributaryArea = CellArea (use actual clipped area) ---
merged_df_c['TributaryArea'] = merged_df_c['CellArea'].copy()

# --- 5) Summary: Storey | ColumnBay | TributaryArea (only) ---
if 'Storey' not in merged_df_c.columns:
    floor_ids_sorted = np.sort(merged_df_c['_floor_id_'].unique())
    storey_map = {fid: i+1 for i, fid in enumerate(floor_ids_sorted)}
    merged_df_c['Storey'] = merged_df_c['_floor_id_'].map(storey_map)

label_col = 'ColumnBay' if 'ColumnBay' in merged_df_c.columns else None
if label_col is None:
    for cand in ['ColumnLabel','ColumnID','ColumnName','UniquePtI','UniquePtJ']:
        if cand in merged_df_c.columns:
            label_col = cand; break
if label_col is None:
    label_col = merged_df_c.columns[0]

summaryE = (
    merged_df_c
    .assign(Column=merged_df_c[label_col].astype(str))
    [['Storey','Column','TributaryArea']]
    .sort_values(['Storey','Column'], kind='mergesort')
    .reset_index(drop=True)
)

summaryE['Storey'] = summaryE['Storey'].apply(lambda x: f"S{x}")
print(summaryE.to_string(index=False, float_format=lambda v: f"{v:.4f}"))

# redraw figure
plt.draw()


# In[17]:


# Full script: scaled-slab redraw AND Voronoi + CellArea computed on scaled outer-only slabs
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d.art3d import Poly3DCollection
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401
from collections import defaultdict

# shapely for geometry ops
try:
    from shapely.geometry import Polygon, MultiPolygon, box, Point
    from shapely.ops import unary_union
    from shapely import affinity as shp_affinity
except Exception as e:
    raise ImportError("This script requires shapely. Install with `pip install shapely` and re-run.") from e

# -------------------------
# Helper functions
# -------------------------
def _to_key(x):
    if pd.isna(x):
        return None
    return str(x).strip()

def parse_unique_pts(s):
    if pd.isna(s):
        return []
    parts = [p.strip() for p in str(s).split(';')]
    return [_to_key(p) for p in parts if p.strip() != ""]

def build_coord_lookup(A):
    A_loc = A.copy()
    A_loc['__key__'] = A_loc['UniqueName'].apply(_to_key)
    return {
        k: (row['X'], row['Y'], row['Z'])
        for k, row in A_loc.set_index('__key__')[['X','Y','Z']].iterrows()
    }

def coords_for_ids(ids, coord_lu):
    out = []
    for pid in ids:
        if pid in coord_lu:
            out.append(coord_lu[pid])
    return out

def set_axes_equal(ax):
    x_limits = ax.get_xlim3d()
    y_limits = ax.get_ylim3d()
    z_limits = ax.get_zlim3d()
    x_range = abs(x_limits[1] - x_limits[0])
    y_range = abs(y_limits[1] - y_limits[0])
    z_range = abs(z_limits[1] - z_limits[0])
    plot_radius = 0.5 * max([x_range, y_range, z_range])
    x_middle = np.mean(x_limits)
    y_middle = np.mean(y_limits)
    z_middle = np.mean(z_limits)
    ax.set_xlim3d([x_middle - plot_radius, x_middle + plot_radius])
    ax.set_ylim3d([y_middle - plot_radius, y_middle + plot_radius])
    ax.set_zlim3d([z_middle - plot_radius, z_middle + plot_radius])

def clip_halfplane(poly, m, n):
    if not poly:
        return []
    out = []
    def inside(p):
        return ((p[0]-m[0])*n[0] + (p[1]-m[1])*n[1]) <= 1e-9
    def intersect(p1, p2):
        dp = (p2[0]-p1[0], p2[1]-p1[1])
        num = (m[0]-p1[0])*n[0] + (m[1]-p1[1])*n[1]
        den = dp[0]*n[0] + dp[1]*n[1]
        if abs(den) < 1e-12:
            return p2
        t = num/den
        return (p1[0]+t*dp[0], p1[1]+t*dp[1])
    S = poly[-1]
    Sin = inside(S)
    for E in poly:
        Ein = inside(E)
        if Ein:
            if not Sin:
                out.append(intersect(S, E))
            out.append(E)
        else:
            if Sin:
                out.append(intersect(S, E))
        S, Sin = E, Ein
    return out

def convex_hull(points_xy):
    pts = np.unique(np.asarray(points_xy, dtype=float), axis=0)
    if len(pts) <= 2:
        return pts.tolist()
    pts = pts[np.lexsort((pts[:,1], pts[:,0]))]
    def cross(o, a, b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    lower = []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(tuple(p))
    upper = []
    for p in pts[::-1]:
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(tuple(p))
    return lower[:-1] + upper[:-1]

# -------------------------
# Validate required DataFrames in environment
# -------------------------
needed_A = {'UniqueName','X','Y','Z'}
if not needed_A.issubset(set(A.columns)):
    raise KeyError(f"A must contain columns {needed_A}; currently has {list(A.columns)}")
for df in (B, C):
    if 'UniquePtI' not in df.columns or 'UniquePtJ' not in df.columns:
        raise KeyError("B and C must each have 'UniquePtI' and 'UniquePtJ' columns.")
if 'UniquePts' not in D.columns:
    raise KeyError("D must have a 'UniquePts' column with 'id1; id2; ...' strings.")
if 'merged_df_c' not in globals():
    raise KeyError("This script expects a DataFrame named merged_df_c containing column coordinates for columns.")

# -------------------------
# Build lookups and parse D
# -------------------------
coord_lu = build_coord_lookup(A)
for df in (B, C):
    df['__I__'] = df['UniquePtI'].apply(_to_key)
    df['__J__'] = df['UniquePtJ'].apply(_to_key)
D = D.copy()
D['__ids__'] = D['UniquePts'].apply(parse_unique_pts)

cols = merged_df_c.columns.tolist()
def find_pair(suffix1, suffix2):
    x1, y1, z1 = f'X{suffix1}', f'Y{suffix1}', f'Z{suffix1}'
    x2, y2, z2 = f'X{suffix2}', f'Y{suffix2}', f'Z{suffix2}'
    if all(x in cols for x in (x1,y1,z1,x2,y2,z2)):
        return (x1,y1,z1), (x2,y2,z2)
    return None

pairs_to_try = [('_A','_C'), ('_C','_A'), ('_bottom','_top'), ('_top','_bottom'), ('_A','_B'), ('','_top')]
bottom_cols = top_cols = None
for s1, s2 in pairs_to_try:
    found = find_pair(s1, s2)
    if found:
        bottom_cols, top_cols = found
        break
if bottom_cols is None:
    x_cols = [c for c in cols if c.startswith('X_')]
    if len(x_cols) >= 2:
        suf_list = list({c.split('_',1)[1] for c in x_cols})
        if len(suf_list) >= 2:
            s1, s2 = suf_list[0], suf_list[1]
            bottom_cols = (f'X_{s1}', f'Y_{s1}', f'Z_{s1}')
            top_cols    = (f'X_{s2}', f'Y_{s2}', f'Z_{s2}')
if bottom_cols is None or top_cols is None:
    raise KeyError(f"Couldn't auto-detect bottom/top coordinate columns. Available cols: {cols}")

xb_col, yb_col, zb_col = bottom_cols
xt_col, yt_col, zt_col = top_cols

floor_decimals = 3
if '_floor_id_' not in merged_df_c.columns:
    merged_df_c['_floor_id_'] = np.round(merged_df_c[zt_col].to_numpy(), floor_decimals)

# ---------- Create figure ----------
fig = plt.figure(figsize=(18,14))
ax = fig.add_subplot(111, projection='3d')

# ---------- Plot beams ----------
for _, row in B.iterrows():
    pI, pJ = row['__I__'], row['__J__']
    if pI in coord_lu and pJ in coord_lu:
        (x1, y1, z1) = coord_lu[pI]
        (x2, y2, z2) = coord_lu[pJ]
        ax.plot([x1, x2], [y1, y2], [z1, z2], linewidth=1.2)

# ---------- Plot columns ----------
for _, row in C.iterrows():
    pI, pJ = row['__I__'], row['__J__']
    if pI in coord_lu and pJ in coord_lu:
        (x1, y1, z1) = coord_lu[pI]
        (x2, y2, z2) = coord_lu[pJ]
        ax.plot([x1, x2], [y1, y2], [z1, z2], linewidth=1.2)

# ---------- Build accurate per-floor slab geometry (with holes) ----------
floor_polygons = {}
floor_points = defaultdict(list)
floor_zlevel = {}
for _, row in D.iterrows():
    ids = row.get('__ids__', [])
    verts = [coord_lu.get(i) for i in ids if i in coord_lu]
    verts = [v for v in verts if v is not None]
    if len(verts) < 3:
        continue
    xs, ys, zs = zip(*verts)
    fid = float(np.round(np.median(zs), floor_decimals))
    floor_points[fid].extend((float(x), float(y)) for x, y in zip(xs, ys))
    floor_zlevel[fid] = float(np.median(zs))
    try:
        poly = Polygon([(float(x), float(y)) for x, y in zip(xs, ys)])
        if not poly.is_valid:
            poly = poly.buffer(0)
        if poly.is_valid and not poly.is_empty:
            floor_polygons.setdefault(fid, []).append(poly)
    except Exception:
        continue

for fid, polys in list(floor_polygons.items()):
    u = unary_union(polys)
    floor_polygons[fid] = u

for fid, pts in floor_points.items():
    if fid not in floor_polygons or floor_polygons[fid] is None or floor_polygons[fid].is_empty:
        hull = convex_hull(pts)
        if len(hull) >= 3:
            floor_polygons[fid] = Polygon(hull)

# ---------- Build outer-only polygons (use outer envelope = convex hull) and compute areas ----------
outer_polys = {}
areas = {}
for fid, poly in floor_polygons.items():
    if poly is None or poly.is_empty:
        continue

    # 1) Build outer envelope from point cloud (convex hull of all floor points)
    pts = floor_points.get(fid, [])
    hull = convex_hull(pts) if pts else None

    if hull and len(hull) >= 3:
        outer = Polygon(hull)            # full outer envelope (fills cut-outs)
    else:
        # fallback: try to get exterior from unioned polygon(s)
        if isinstance(poly, MultiPolygon):
            parts = list(poly.geoms)
            largest = max(parts, key=lambda p: p.area)
            outer = Polygon(largest.exterior.coords)
        elif isinstance(poly, Polygon):
            outer = Polygon(poly.exterior.coords)
        else:
            continue

    # ensure valid
    if not outer.is_valid:
        outer = outer.buffer(0)
        if not outer.is_valid or outer.is_empty:
            continue

    outer_polys[fid] = outer
    areas[fid] = float(outer.area)

if not areas:
    raise RuntimeError("No slab outer polygons found to scale.")

max_area = max(areas.values())

# ---------- Scale each outer polygon to the maximum area ----------
#scaled_floor_polygons = {}
#for fid, outer in outer_polys.items():
    #a = areas[fid]
    #if a <= 0:
        #scaled_floor_polygons[fid] = outer
        #continue
    #scale_factor = (max_area / a) ** 0.5
    #centroid = outer.representative_point()
    #scaled = shp_affinity.scale(outer, xfact=scale_factor, yfact=scale_factor, origin=(centroid.x, centroid.y))
    #if not scaled.is_valid:
        #scaled = scaled.buffer(0)
    #scaled_floor_polygons[fid] = scaled
    
# ---------- Option A: use the single largest outer polygon as the master footprint ----------
# Find the fid with the maximum outer area and use its outer polygon unchanged for all floors.
fid_master = max(areas, key=areas.get)
master_poly = outer_polys[fid_master]

# sanity check
if master_poly is None or master_poly.is_empty:
    raise RuntimeError(f"Chosen master polygon for fid {fid_master} is empty or invalid.")

# 1) Assign the master outer polygon to every floor that HAS a slab (outer_polys keys)
scaled_floor_polygons = {fid: master_poly for fid in outer_polys.keys()}

# 2) ALSO assign the master polygon to floors with NO slab (but columns exist there)
all_fids = merged_df_c['_floor_id_'].unique()
for fid in all_fids:
    if fid not in scaled_floor_polygons:
        scaled_floor_polygons[fid] = master_poly

print(
    f"Using fid {fid_master} as master footprint (outer area = {areas[fid_master]:.3f}). "
    f"Assigned master footprint to {len(scaled_floor_polygons)} floors "
    f"(including floors with no slab rows in D)."
)


# ---------- Draw scaled slabs (outer-only, holes ignored) ----------
for fid, scaled in scaled_floor_polygons.items():
    z_floor = floor_zlevel.get(fid, None)
    if z_floor is None:
        z_floor_vals = merged_df_c.loc[merged_df_c['_floor_id_'] == fid, zt_col].to_numpy()
        z_floor = float(np.median(z_floor_vals)) if z_floor_vals.size else 0.0
    if scaled.is_empty:
        continue
    parts = list(scaled.geoms) if isinstance(scaled, MultiPolygon) else [scaled]
    for part in parts:
        if part.exterior is None:
            continue
        coords2d = list(part.exterior.coords)
        verts3d = [(x, y, z_floor) for (x, y) in coords2d]
        poly3d = Poly3DCollection([verts3d],
                                  facecolors=(0.85,0.85,0.85,0.85),
                                  edgecolors='k', linewidths=1.2, zorder=1)
        ax.add_collection3d(poly3d)

print(f"Scaled {len(scaled_floor_polygons)} floors to max outer-area = {max_area:.3f} (holes ignored).")

# ---------- Draw columns and star markers at original coordinates ----------
mean_zb = merged_df_c[zb_col].mean()
mean_zt = merged_df_c[zt_col].mean()
if mean_zt < mean_zb:
    xb_col, xt_col = xt_col, xb_col
    yb_col, yt_col = yt_col, yb_col
    zb_col, zt_col = zt_col, zb_col

for i, row in merged_df_c.iterrows():
    x = [row[xb_col], row[xt_col]]
    y = [row[yb_col], row[yt_col]]
    z = [row[zb_col], row[zt_col]]
    ax.plot(x, y, z, color='black', linewidth=1.0, label='Column' if i == 0 else '')
    offset = 0.05
    ax.scatter(row[xt_col], row[yt_col], row[zt_col] + offset,
               color='red', marker='*', s=200, edgecolor='k', linewidth=0.8, zorder=10)

# ---------- Dark-blue neighbor lines on each floor ----------
merged_df_c['_floor_id_'] = np.round(merged_df_c[zt_col].to_numpy(), floor_decimals)
k_nearest = None
max_radius = None
_drawn_pairs = set()
for fid, df_floor in merged_df_c.groupby('_floor_id_'):
    if df_floor.shape[0] < 2:
        continue
    idxs = df_floor.index.to_list()
    XY = df_floor[[xt_col, yt_col]].to_numpy(dtype=float)
    Z  = df_floor[zt_col].to_numpy(dtype=float)
    for i, idx_i in enumerate(idxs):
        x0, y0, z0 = XY[i,0], XY[i,1], Z[i]
        d2 = np.sum((XY - XY[i])**2, axis=1)
        d  = np.sqrt(d2)
        order = np.argsort(d)
        order = [j for j in order if j != i]
        if max_radius is not None:
            order = [j for j in order if d[j] <= max_radius]
        if k_nearest is not None:
            order = order[:k_nearest]
        if not order:
            continue
        for j in order:
            idx_j = idxs[j]
            pairkey = frozenset((idx_i, idx_j))
            if pairkey in _drawn_pairs:
                continue
            _drawn_pairs.add(pairkey)
            x1, y1, z1 = XY[j,0], XY[j,1], Z[j]
            ax.plot([x0, x1], [y0, y1], [z0, z1], color='#003366', linewidth=3.0, linestyle=':', zorder=5)
            xm = (x0 + x1) / 2.0
            ym = (y0 + y1) / 2.0
            zm = (z0 + z1) / 2.0
            ax.scatter(xm, ym, zm, color='black', s=30, zorder=8)
            dx = x1 - x0
            dy = y1 - y0
            seg_len = (dx*dx + dy*dy) ** 0.5
            if seg_len > 0:
                px = -dy / seg_len
                py =  dx / seg_len
                half_len = 0.25 * seg_len
                xL, yL = xm - px * half_len, ym - py * half_len
                xR, yR = xm + px * half_len, ym + py * half_len
                ax.plot([xL, xR], [yL, yR], [zm, zm], color='#FF8C00', linewidth=2.5, zorder=9)

# ---------- SHADE Voronoi-like cells per floor, CLIPPED TO SCALED (outer-only) SLABS ----------
from matplotlib import colormaps
cmap = colormaps.get_cmap('tab20')

for fid, df_floor in merged_df_c.groupby('_floor_id_'):
    if df_floor.shape[0] == 0:
        continue
    XY = df_floor[[xt_col, yt_col]].to_numpy(dtype=float)
    Z  = df_floor[zt_col].to_numpy(dtype=float)
    z_floor = float(np.median(Z)) if fid not in floor_zlevel else floor_zlevel[fid]

    # <-- USE SCALED polygon here (holes ignored)
    slab_poly = scaled_floor_polygons.get(fid, None)
    if slab_poly is None or slab_poly.is_empty:
        minx, miny = XY.min(axis=0); maxx, maxy = XY.max(axis=0)
        padx = 0.1*(maxx-minx if maxx>minx else 1.0)
        pady = 0.1*(maxy-miny if maxy>miny else 1.0)
        slab_poly = box(minx-padx, miny-pady, maxx+padx, maxy+pady)

    ncols = len(df_floor)
    for i in range(ncols):
        si = XY[i]
        minx, miny, maxx, maxy = slab_poly.bounds
        padx = 0.2*(maxx-minx if maxx>minx else 1.0)
        pady = 0.2*(maxy-miny if maxy>miny else 1.0)
        bounds_poly = [(minx-padx,miny-pady),(maxx+padx,miny-pady),
                       (maxx+padx,maxy+pady),(minx-padx,maxy+pady)]
        cell = bounds_poly[:]
        for j in range(ncols):
            if j == i:
                continue
            sj = XY[j]
            m = 0.5*(si + sj)
            n = (sj - si)
            cell = clip_halfplane(cell, m, n)
            if not cell:
                break
        if not cell:
            continue
        try:
            cell_poly = Polygon(cell)
            if not cell_poly.is_valid:
                cell_poly = cell_poly.buffer(0)
            clipped = cell_poly.intersection(slab_poly)
        except Exception:
            continue
        if clipped.is_empty:
            continue
        parts = list(clipped.geoms) if isinstance(clipped, MultiPolygon) else [clipped]
        color = cmap((i % 20)/20.0)
        for part in parts:
            if not part.exterior:
                continue
            coords2d = list(part.exterior.coords)
            verts3d = [(x, y, z_floor) for (x, y) in coords2d]
            poly = Poly3DCollection([verts3d], facecolors=color, edgecolors='k',
                                    linewidths=0.6, alpha=0.35, zorder=3)
            ax.add_collection3d(poly)

# ---------- Compute CellArea and TributaryArea USING SCALED (outer-only) SLABS ----------
merged_df_c['CellArea'] = np.nan

for fid, df_floor in merged_df_c.groupby('_floor_id_'):
    XY = df_floor[[xt_col, yt_col]].to_numpy(dtype=float)
    slab_poly = scaled_floor_polygons.get(fid, None)
    if slab_poly is None or slab_poly.is_empty:
        if XY.shape[0] == 0:
            continue
        minx, miny = XY.min(axis=0); maxx, maxy = XY.max(axis=0)
        padx = 0.1*(maxx-minx if maxx>minx else 1.0)
        pady = 0.1*(maxy-miny if maxy>miny else 1.0)
        slab_poly = box(minx-padx, miny-pady, maxx+padx, maxy+pady)

    idxs = df_floor.index.to_list()
    ncols = len(df_floor)
    for i in range(ncols):
        si = XY[i]
        minx, miny, maxx, maxy = slab_poly.bounds
        padx = 0.2*(maxx-minx if maxx>minx else 1.0)
        pady = 0.2*(maxy-miny if maxy>miny else 1.0)
        bounds_poly = [(minx-padx,miny-pady),(maxx+padx,miny-pady),
                       (maxx+padx,maxy+pady),(minx-padx,maxy+pady)]
        cell = bounds_poly[:]
        for j in range(ncols):
            if j == i:
                continue
            sj = XY[j]
            m = 0.5*(si + sj)
            n = (sj - si)
            cell = clip_halfplane(cell, m, n)
            if not cell:
                break
        if not cell:
            area = 0.0
        else:
            try:
                cell_poly = Polygon(cell)
                if not cell_poly.is_valid:
                    cell_poly = cell_poly.buffer(0)
                clipped = cell_poly.intersection(slab_poly)
                area = float(clipped.area) if not clipped.is_empty else 0.0
            except Exception:
                area = 0.0
        merged_df_c.loc[idxs[i], 'CellArea'] = area

merged_df_c['TributaryArea'] = merged_df_c['CellArea'].copy()

# ---------- Summary printout (note: computed on scaled outer-only slabs) ----------
if 'Storey' not in merged_df_c.columns:
    floor_ids_sorted = np.sort(merged_df_c['_floor_id_'].unique())
    storey_map = {fid: i+1 for i, fid in enumerate(floor_ids_sorted)}
    merged_df_c['Storey'] = merged_df_c['_floor_id_'].map(storey_map)

label_col = 'ColumnBay' if 'ColumnBay' in merged_df_c.columns else None
if label_col is None:
    for cand in ['ColumnLabel','ColumnID','ColumnName','UniquePtI','UniquePtJ']:
        if cand in merged_df_c.columns:
            label_col = cand; break
if label_col is None:
    label_col = merged_df_c.columns[0]

summary = (
    merged_df_c
    .assign(Column=merged_df_c[label_col].astype(str))
    [['Storey','Column','TributaryArea']]
    .sort_values(['Storey','Column'], kind='mergesort')
    .reset_index(drop=True)
)
summary['Storey'] = summary['Storey'].apply(lambda x: f"S{x}")
pd.options.display.float_format = '{:,.4f}'.format
print("\nTributary area summary (TributaryArea computed on SCALED outer-only slabs):\n")
print(summary.to_string(index=False))

# ---------- Cosmetics and show ----------
ax.set_xlabel('X')
ax.set_ylabel('Y')
ax.set_zlabel('Z')
ax.set_title('3D Frame: scaled slabs (outer-only) + columns + Voronoi-like cells (clipped to SCALED slabs)')
set_axes_equal(ax)
plt.tight_layout()
plt.show()


# In[18]:


table = model.DatabaseTables.GetTableForDisplayArray("Element Forces - Columns", "", "")
cols = table[2]
noOfRows = table[3]
vals = np.array_split(table[4], noOfRows)
CF = pd.DataFrame(vals, columns=cols)
#print(CF)

CF = CF[CF["OutputCase"] == "1.5DL+1.5SDL+1.5LL"]
#print(CF)

#print("Columns returned by ETABS:", CF.columns.tolist())

# Convert 'Station m' to numeric
CF["Station"] = pd.to_numeric(CF["Station"], errors="coerce")

# Keep only bottom-most station (Station m == 0)
CF = CF[CF["Station"] == CF["Station"].min()].copy()

# Optional: reset index for cleaner display
CF.reset_index(drop=True, inplace=True)




columns_to_remove = ["UniqueName", "CaseType", "Station", "V2", "V3", "T", "M2", "M3", "Element", "ElemStation", "Location"]
# Drop the specified columns from the DataFrame
CF = CF.drop(columns=columns_to_remove)
#print(CF)


CF['Story'] = (
    CF['Story']
    .astype(str)
    .str.strip()
    .str.extract(r'(\d+)')[0]     # <-- Pick the first column returned by extract
    .apply(lambda x: f"S{x}" if pd.notnull(x) else None)
)
#print(CF)
#display(CF)


CF['FloorNum'] = CF['Story'].str.extract(r'(\d+)').astype(int)

# Position from top: S5->1, S4->2, ..., S1->5 (for 5 floors)
max_floor = CF['FloorNum'].max()
CF['n'] = max_floor - CF['FloorNum'] + 1

CF.drop(columns=['FloorNum'], inplace=True)
CF = CF.rename(columns={'Story': 'Storey'})
#print(CF.to_string(index=False))
#print(CF)




# 0) normalize keys (avoids hidden mismatches like spaces/case)
for df in (CF, summary):
    for k in ['Storey','Column']:
        df[k] = df[k].astype(str).str.strip()

# 1) remove any previous TributaryArea columns in CF (including _x/_y)
trib_cols = [c for c in CF.columns if c.startswith('TributaryArea')]
CF = CF.drop(columns=trib_cols, errors='ignore')

# 2) keep only one row per (Storey, Column) in summary
summary_u = summary.drop_duplicates(['Storey','Column'])[
    ['Storey','Column','TributaryArea']
]

# 3) merge in the area (this cannot create _x/_y now)
CF = CF.merge(summary_u, on=['Storey','Column'], how='left')

# 4) (optional) print without index, with 4 decimals
#print(CF.to_string(index=False, float_format=lambda v: f"{v:.4f}"))

CF['P'] = pd.to_numeric(CF['P'], errors='coerce')
CF['TributaryArea'] = pd.to_numeric(CF['TributaryArea'], errors='coerce')
CF['UIF'] = CF['P'] / (CF['TributaryArea'] * CF['n'] * (- 1000))



print(CF.to_string(index=False, float_format=lambda v: f"{v:.4f}"))


# In[19]:


# Plotly interactive replacement
import numpy as np
import plotly.graph_objects as go

# re-use your helpers: build_coord_lookup, parse_unique_pts, coords_for_ids
coord_lu = build_coord_lookup(A)
for df in (B, C):
    df['__I__'] = df['UniquePtI'].apply(_to_key)
    df['__J__'] = df['UniquePtJ'].apply(_to_key)
D = D.copy()
D['__ids__'] = D['UniquePts'].apply(parse_unique_pts)

# Build line traces for beams+columns (single trace each for performance)
def build_line_trace(df, coord_lu, name):
    xs, ys, zs = [], [], []
    for _, row in df.iterrows():
        pI, pJ = row['__I__'], row['__J__']
        if pI in coord_lu and pJ in coord_lu:
            x1,y1,z1 = coord_lu[pI]; x2,y2,z2 = coord_lu[pJ]
            xs += [x1, x2, None]   # None creates breaks between segments
            ys += [y1, y2, None]
            zs += [z1, z2, None]
    return go.Scatter3d(x=xs, y=ys, z=zs, mode='lines', name=name, line=dict(width=4)
                       )

beam_trace = build_line_trace(B, coord_lu, 'Beams')
beam_trace.update(hoverinfo='skip')

col_trace  = build_line_trace(C, coord_lu, 'Columns')
col_trace.update(hoverinfo='skip')

# Nodes (optional)
nx, ny, nz, nlabels = [], [], [], []
for k,(x,y,z) in coord_lu.items():
    nx.append(x); ny.append(y); nz.append(z); nlabels.append(str(k))
nodes = go.Scatter3d(x=nx, y=ny, z=nz, mode='markers', name='Nodes',
                     marker=dict(size=3), text=nlabels, hoverinfo='skip')

# Slabs: convert each polygon to triangles by fan triangulation (good for convex/planar)
mesh_traces = []
for _, row in D.iterrows():
    ids = row['__ids__']; verts = coords_for_ids(ids, coord_lu)
    if len(verts) < 3: 
        continue
    verts = np.array(verts)
    x, y, z = verts[:,0], verts[:,1], verts[:,2]

    # Fan triangulation: triangles (0,i,i+1) for i=1..n-2
    i_idxs = []; j_idxs = []; k_idxs = []
    for i in range(1, len(verts)-1):
        i_idxs.append(0); j_idxs.append(i); k_idxs.append(i+1)

    mesh = go.Mesh3d(x=x, y=y, z=z,
                     i=i_idxs, j=j_idxs, k=k_idxs,
                     opacity=0.6, name=f"Slab {row.get('FloorBay','')}",
                     hoverinfo='skip')
    mesh_traces.append(mesh)
    
for mesh in mesh_traces:
    mesh.color = 'lightgreen'   # slab color
    

fig = go.Figure(data=[beam_trace, col_trace, nodes] + mesh_traces)
fig.update_layout(scene=dict(
    xaxis=dict(showgrid=False, showticklabels=False),  # Remove gridlines and ticks from X axis
    yaxis=dict(showgrid=False, showticklabels=False),  # Remove gridlines and ticks from Y axis
    zaxis=dict(showgrid=False, showticklabels=False),  # Remove gridlines and ticks from Z axis
    aspectmode='data'  # Preserve aspect ratio
), title='Interactive 3D Frame (Plotly)')






fig.show()


# In[20]:


get_ipython().run_cell_magic('capture', '', '# plotly_scaled_slab.py - minimalist version (no grid, no axis labels, only hovers)\n\nimport pandas as pd\nimport numpy as np\nimport plotly.graph_objects as go\nfrom shapely.geometry import Polygon, MultiPolygon, box, Point\nfrom shapely.ops import unary_union\nfrom shapely import affinity as shp_affinity\n\n# (All helper functions and geometry computations unchanged from previous version)\n# Assume A, B, C, D, merged_df_c are already prepared in memory.\n\n# --- Only plotting section below for brevity ---\n\nfig = go.Figure()\n\n\n# 1) Add beams (B) as 3D lines\nfor _, row in B.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y1, y2], z=[z1, z2], mode=\'lines\',\n                                   line=dict(width=4, color=\'blue\'), hoverinfo=\'none\', name=\'\'))\n        \n\n# Add columns as 3D lines\nfor _, row in C.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y1, y2], z=[z1, z2], mode=\'lines\',\n                                   line=dict(width=6, color=\'red\'), hoverinfo=\'none\', name=\'\'))\n        \n\n# Add slabs (original, unscaled)\nfor fid, poly in floor_polygons.items():\n    if poly is None or poly.is_empty:\n        continue\n    parts = list(poly.geoms) if isinstance(poly, MultiPolygon) else [poly]\n    z_floor = floor_zlevel.get(fid, 0.0)\n    for part in parts:\n        if part.exterior is None:\n            continue\n        coords = list(part.exterior.coords)\n        if len(coords) >= 3:\n            xs, ys = zip(*coords)\n            zs = [z_floor]*len(xs)\n            fig.add_trace(go.Mesh3d(x=xs, y=ys, z=zs, color=\'lightgreen\', opacity=0.9, showscale=False, hoverinfo=\'none\', name=\'\'))\n\n# Hover markers at mid-height\nhover_x, hover_y, hover_z, hover_text = [], [], [], []\nfor _, row in merged_df_c.iterrows():\n    hx = row[xt_col]; hy = row[yt_col]\n    hz = (row[zt_col] + row[zb_col]) / 2\n    txt = f"Storey: {row[\'Storey\']}<br>Column: {row[label_col]}<br>TributaryArea: {row[\'TributaryArea\']:.4f}"\n    hover_x.append(hx); hover_y.append(hy); hover_z.append(hz); hover_text.append(txt)\n\nfig.add_trace(go.Scatter3d(x=hover_x, y=hover_y, z=hover_z,\n                           mode=\'markers\', marker=dict(size=2, color=\'red\'),\n                           hovertext=hover_text, hoverinfo=\'text\', name=\'\'))\n\n# Minimal layout: no grid, no axes, no labels\nfig.update_layout(\n    scene=dict(\n        xaxis=dict(visible=False),\n        yaxis=dict(visible=False),\n        zaxis=dict(visible=False),\n        bgcolor=\'white\'\n    ),\n    paper_bgcolor=\'white\',\n    plot_bgcolor=\'white\',\n    showlegend=False,\n    margin=dict(l=0, r=0, t=0, b=0)\n)\n\nfig.update_layout(scene=dict(aspectmode=\'data\'))\nfig.show()\n')


# In[21]:


get_ipython().run_cell_magic('capture', '', 'print(merged_df_c)\n')


# In[22]:


# Make S# key on the left
merged_df_c['Storey_str'] = merged_df_c['Storey'].apply(lambda x: f"S{int(x)}")

# Take only what we need from CF and rename keys to match left
cf_small = CF[['Storey', 'Column', 'P']].copy()
cf_small = cf_small.rename(columns={'Storey': 'Storey_str', 'Column': 'ColumnBay'})

# Merge on the aligned keys
merged_df_c = merged_df_c.merge(cf_small, on=['Storey_str', 'ColumnBay'], how='left')

# Tidy up
merged_df_c = merged_df_c.drop(columns=['Storey_str'])
# (optional) remove duplicate rows if any
merged_df_c = merged_df_c.drop_duplicates(subset=['Storey', 'ColumnBay'], keep='last')


# In[23]:


get_ipython().run_cell_magic('capture', '', 'merged_df_c.columns.tolist()\n')


# In[24]:


get_ipython().run_cell_magic('capture', '', 'print(merged_df_c)\n')


# In[25]:


get_ipython().run_cell_magic('capture', '', '# plotly_scaled_slab.py - minimalist version (no grid, no axis labels, only hovers)\n\nimport pandas as pd\nimport numpy as np\nimport plotly.graph_objects as go\nfrom shapely.geometry import Polygon, MultiPolygon, box, Point\nfrom shapely.ops import unary_union\nfrom shapely import affinity as shp_affinity\n\n# (All helper functions and geometry computations unchanged from previous version)\n# Assume A, B, C, D, merged_df_c are already prepared in memory.\n\n# --- Only plotting section below for brevity ---\n\nfig = go.Figure()\n\n\n# 1) Add beams (B) as 3D lines\nfor _, row in B.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y1, y2], z=[z1, z2], mode=\'lines\',\n                                   line=dict(width=4, color=\'blue\'), hoverinfo=\'none\', name=\'\'))\n        \n\n# Add columns as 3D lines\nfor _, row in C.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y1, y2], z=[z1, z2], mode=\'lines\',\n                                   line=dict(width=6, color=\'red\'), hoverinfo=\'none\', name=\'\'))\n        \n\n# Add slabs (original, unscaled)\nfor fid, poly in floor_polygons.items():\n    if poly is None or poly.is_empty:\n        continue\n    parts = list(poly.geoms) if isinstance(poly, MultiPolygon) else [poly]\n    z_floor = floor_zlevel.get(fid, 0.0)\n    for part in parts:\n        if part.exterior is None:\n            continue\n        coords = list(part.exterior.coords)\n        if len(coords) >= 3:\n            xs, ys = zip(*coords)\n            zs = [z_floor]*len(xs)\n            fig.add_trace(go.Mesh3d(x=xs, y=ys, z=zs, color=\'lightgreen\', opacity=0.9, showscale=False, hoverinfo=\'none\', name=\'\'))\n\n# Hover markers at mid-height\nhover_x, hover_y, hover_z, hover_text = [], [], [], []\nfor _, row in merged_df_c.iterrows():\n    hx = row[xt_col]; hy = row[yt_col]\n    hz = (row[zt_col] + row[zb_col]) / 2\n    txt = f"Storey: {row[\'Storey\']}<br>Column: {row[label_col]}<br>P: {row[\'P\']:.4f}"\n    hover_x.append(hx); hover_y.append(hy); hover_z.append(hz); hover_text.append(txt)\n\nfig.add_trace(go.Scatter3d(x=hover_x, y=hover_y, z=hover_z,\n                           mode=\'markers\', marker=dict(size=2, color=\'red\'),\n                           hovertext=hover_text, hoverinfo=\'text\', name=\'\'))\n\n# Minimal layout: no grid, no axes, no labels\nfig.update_layout(\n    scene=dict(\n        xaxis=dict(visible=False),\n        yaxis=dict(visible=False),\n        zaxis=dict(visible=False),\n        bgcolor=\'white\'\n    ),\n    paper_bgcolor=\'white\',\n    plot_bgcolor=\'white\',\n    showlegend=False,\n    margin=dict(l=0, r=0, t=0, b=0)\n)\n\nfig.update_layout(scene=dict(aspectmode=\'data\'))\nfig.show()\n')


# In[26]:


# Make S# key on the left
merged_df_c['Storey_str'] = merged_df_c['Storey'].apply(lambda x: f"S{int(x)}")

# Take only what we need from CF and rename keys to match left
cf_small = CF[['Storey', 'Column', 'UIF']].copy()
cf_small = cf_small.rename(columns={'Storey': 'Storey_str', 'Column': 'ColumnBay'})

# Merge on the aligned keys
merged_df_c = merged_df_c.merge(cf_small, on=['Storey_str', 'ColumnBay'], how='left')

# Tidy up
merged_df_c = merged_df_c.drop(columns=['Storey_str'])
# (optional) remove duplicate rows if any
merged_df_c = merged_df_c.drop_duplicates(subset=['Storey', 'ColumnBay'], keep='last')


# In[27]:


get_ipython().run_cell_magic('capture', '', 'merged_df_c.columns.tolist()\n')


# In[28]:


get_ipython().run_cell_magic('capture', '', '# plotly_scaled_slab.py - minimalist version (no grid, no axis labels, only hovers)\n\nimport pandas as pd\nimport numpy as np\nimport plotly.graph_objects as go\nfrom shapely.geometry import Polygon, MultiPolygon, box, Point\nfrom shapely.ops import unary_union\nfrom shapely import affinity as shp_affinity\n\n# (All helper functions and geometry computations unchanged from previous version)\n# Assume A, B, C, D, merged_df_c are already prepared in memory.\n\n# --- Only plotting section below for brevity ---\n\nfig = go.Figure()\n\n\n# 1) Add beams (B) as 3D lines\nfor _, row in B.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y1, y2], z=[z1, z2], mode=\'lines\',\n                                   line=dict(width=4, color=\'blue\'), hoverinfo=\'none\', name=\'\'))\n        \n\n# Add columns as 3D lines\nfor _, row in C.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y1, y2], z=[z1, z2], mode=\'lines\',\n                                   line=dict(width=6, color=\'red\'), hoverinfo=\'none\', name=\'\'))\n        \n\n# Add slabs (original, unscaled)\nfor fid, poly in floor_polygons.items():\n    if poly is None or poly.is_empty:\n        continue\n    parts = list(poly.geoms) if isinstance(poly, MultiPolygon) else [poly]\n    z_floor = floor_zlevel.get(fid, 0.0)\n    for part in parts:\n        if part.exterior is None:\n            continue\n        coords = list(part.exterior.coords)\n        if len(coords) >= 3:\n            xs, ys = zip(*coords)\n            zs = [z_floor]*len(xs)\n            fig.add_trace(go.Mesh3d(x=xs, y=ys, z=zs, color=\'lightgreen\', opacity=0.9, showscale=False, hoverinfo=\'none\', name=\'\'))\n\n# Hover markers at mid-height\nhover_x, hover_y, hover_z, hover_text = [], [], [], []\nfor _, row in merged_df_c.iterrows():\n    hx = row[xt_col]; hy = row[yt_col]\n    hz = (row[zt_col] + row[zb_col]) / 2\n    txt = f"Storey: {row[\'Storey\']}<br>Column: {row[label_col]}<br>UIF: {row[\'UIF\']:.4f}"\n    hover_x.append(hx); hover_y.append(hy); hover_z.append(hz); hover_text.append(txt)\n\nfig.add_trace(go.Scatter3d(x=hover_x, y=hover_y, z=hover_z,\n                           mode=\'markers\', marker=dict(size=2, color=\'red\'),\n                           hovertext=hover_text, hoverinfo=\'text\', name=\'\'))\n\n# Minimal layout: no grid, no axes, no labels\nfig.update_layout(\n    scene=dict(\n        xaxis=dict(visible=False),\n        yaxis=dict(visible=False),\n        zaxis=dict(visible=False),\n        bgcolor=\'white\'\n    ),\n    paper_bgcolor=\'white\',\n    plot_bgcolor=\'white\',\n    showlegend=False,\n    margin=dict(l=0, r=0, t=0, b=0)\n)\n\n\n# Set reasonable camera (optional)\nfig.update_layout(scene_camera=dict(eye=dict(x=1.25, y=1.25, z=0.8)))\n\n\n# ---------- Add XYZ axes with arrowheads ----------\naxis_len = 1  # adjust based on your model scale\narrow_size = 0.2 * axis_len  # arrowhead size\n\n# Axis lines\nfig.add_trace(go.Scatter3d(\n    x=[0, axis_len], y=[0, 0], z=[0, 0],\n    mode=\'lines\',\n    line=dict(color=\'red\', width=10),\n    hoverinfo=\'none\',\n    showlegend=False\n))\nfig.add_trace(go.Scatter3d(\n    x=[0, 0], y=[0, axis_len], z=[0, 0],\n    mode=\'lines\',\n    line=dict(color=\'green\', width=10),\n    hoverinfo=\'none\',\n    showlegend=False\n))\nfig.add_trace(go.Scatter3d(\n    x=[0, 0], y=[0, 0], z=[0, axis_len],\n    mode=\'lines\',\n    line=dict(color=\'blue\', width=10),\n    hoverinfo=\'none\',\n    showlegend=False\n))\n\n# Arrowheads using cones\nfig.add_trace(go.Cone(\n    x=[axis_len], y=[0], z=[0],\n    u=[1], v=[0], w=[0],\n    sizemode="absolute",\n    sizeref=arrow_size,\n    colorscale=[[0, \'red\'], [1, \'red\']],\n    showscale=False,\n    anchor="tail",\n    hoverinfo="skip"\n))\nfig.add_trace(go.Cone(\n    x=[0], y=[axis_len], z=[0],\n    u=[0], v=[1], w=[0],\n    sizemode="absolute",\n    sizeref=arrow_size,\n    colorscale=[[0, \'green\'], [1, \'green\']],\n    showscale=False,\n    anchor="tail",\n    hoverinfo="skip"\n))\nfig.add_trace(go.Cone(\n    x=[0], y=[0], z=[axis_len],\n    u=[0], v=[0], w=[1],\n    sizemode="absolute",\n    sizeref=arrow_size,\n    colorscale=[[0, \'blue\'], [1, \'blue\']],\n    showscale=False,\n    anchor="tail",\n    hoverinfo="skip"\n))\n\n# Axis labels\nfig.add_trace(go.Scatter3d(\n    x=[axis_len*1.1, 0, 0],\n    y=[0, axis_len*1.1, 0],\n    z=[0, 0, axis_len*1.1],\n    mode=\'text\',\n    text=[\'X\', \'Y\', \'Z\'],\n    textfont=dict(size=18, color=[\'red\',\'green\',\'blue\']),\n    hoverinfo=\'none\',\n    showlegend=False\n))\n\n\n\nfig.update_layout(scene=dict(aspectmode=\'data\'))\nfig.show()\n')


# In[29]:


get_ipython().run_cell_magic('capture', '', '# plotly_scaled_slab.py - minimalist version (no grid, no axis labels, only hovers)\n\nimport pandas as pd\nimport numpy as np\nimport plotly.graph_objects as go\nfrom shapely.geometry import Polygon, MultiPolygon, box, Point\nfrom shapely.ops import unary_union\nfrom shapely import affinity as shp_affinity\n\n# (All helper functions and geometry computations unchanged from previous version)\n# Assume A, B, C, D, merged_df_c are already prepared in memory.\n\n# --- Only plotting section below for brevity ---\n\nfig = go.Figure()\n\n\n# 1) Add beams (B) as 3D lines\nfor _, row in B.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y1, y2], z=[z1, z2], mode=\'lines\',\n                                   line=dict(width=4, color=\'blue\'), hoverinfo=\'none\', name=\'\'))\n        \n\n# Add columns as 3D lines\nfor _, row in C.iterrows():\n    pI, pJ = row[\'__I__\'], row[\'__J__\']\n    if pI in coord_lu and pJ in coord_lu:\n        (x1, y1, z1) = coord_lu[pI]\n        (x2, y2, z2) = coord_lu[pJ]\n        fig.add_trace(go.Scatter3d(x=[x1, x2], y=[y1, y2], z=[z1, z2], mode=\'lines\',\n                                   line=dict(width=6, color=\'grey\'), hoverinfo=\'none\', name=\'\'))\n        \n\n# Add slabs (original, unscaled)\nfor fid, poly in floor_polygons.items():\n    if poly is None or poly.is_empty:\n        continue\n    parts = list(poly.geoms) if isinstance(poly, MultiPolygon) else [poly]\n    z_floor = floor_zlevel.get(fid, 0.0)\n    for part in parts:\n        if part.exterior is None:\n            continue\n        coords = list(part.exterior.coords)\n        if len(coords) >= 3:\n            xs, ys = zip(*coords)\n            zs = [z_floor]*len(xs)\n            fig.add_trace(go.Mesh3d(x=xs, y=ys, z=zs, color=\'lightgreen\', opacity=0.9, showscale=False, hoverinfo=\'none\', name=\'\'))\n\n# Hover markers at mid-height\nhover_x, hover_y, hover_z, hover_text = [], [], [], []\nfor _, row in merged_df_c.iterrows():\n    hx = row[xt_col]; hy = row[yt_col]\n    hz = (row[zt_col] + row[zb_col]) / 2\n    txt = f"Storey: {row[\'Storey\']}<br>Column: {row[label_col]}<br>UIF: {row[\'UIF\']:.2f}"\n    hover_x.append(hx); hover_y.append(hy); hover_z.append(hz); hover_text.append(txt)\n\nfig.add_trace(go.Scatter3d(x=hover_x, y=hover_y, z=hover_z,\n                           mode=\'markers\', marker=dict(size=2, color=\'white\'),\n                           hovertext=hover_text, hoverinfo=\'text\', name=\'\'))\n\n# Minimal layout: no grid, no axes, no labels\nfig.update_layout(\n    scene=dict(\n        xaxis=dict(visible=False),\n        yaxis=dict(visible=False),\n        zaxis=dict(visible=False),\n        bgcolor=\'white\'\n    ),\n    paper_bgcolor=\'white\',\n    plot_bgcolor=\'white\',\n    showlegend=False,\n    margin=dict(l=0, r=0, t=0, b=0)\n)\n\n\n# Set reasonable camera (optional)\nfig.update_layout(scene_camera=dict(eye=dict(x=1.25, y=1.25, z=0.8)))\n\n\n# ---------- Add XYZ axes with arrowheads ----------\naxis_len = 1  # adjust based on your model scale\narrow_size = 0.2 * axis_len  # arrowhead size\n\n# Axis lines\nfig.add_trace(go.Scatter3d(\n    x=[0, axis_len], y=[0, 0], z=[0, 0],\n    mode=\'lines\',\n    line=dict(color=\'red\', width=10),\n    hoverinfo=\'none\',\n    showlegend=False\n))\nfig.add_trace(go.Scatter3d(\n    x=[0, 0], y=[0, axis_len], z=[0, 0],\n    mode=\'lines\',\n    line=dict(color=\'green\', width=10),\n    hoverinfo=\'none\',\n    showlegend=False\n))\nfig.add_trace(go.Scatter3d(\n    x=[0, 0], y=[0, 0], z=[0, axis_len],\n    mode=\'lines\',\n    line=dict(color=\'blue\', width=10),\n    hoverinfo=\'none\',\n    showlegend=False\n))\n\n# Arrowheads using cones\nfig.add_trace(go.Cone(\n    x=[axis_len], y=[0], z=[0],\n    u=[1], v=[0], w=[0],\n    sizemode="absolute",\n    sizeref=arrow_size,\n    colorscale=[[0, \'red\'], [1, \'red\']],\n    showscale=False,\n    anchor="tail",\n    hoverinfo="skip"\n))\nfig.add_trace(go.Cone(\n    x=[0], y=[axis_len], z=[0],\n    u=[0], v=[1], w=[0],\n    sizemode="absolute",\n    sizeref=arrow_size,\n    colorscale=[[0, \'green\'], [1, \'green\']],\n    showscale=False,\n    anchor="tail",\n    hoverinfo="skip"\n))\nfig.add_trace(go.Cone(\n    x=[0], y=[0], z=[axis_len],\n    u=[0], v=[0], w=[1],\n    sizemode="absolute",\n    sizeref=arrow_size,\n    colorscale=[[0, \'blue\'], [1, \'blue\']],\n    showscale=False,\n    anchor="tail",\n    hoverinfo="skip"\n))\n\n# Axis labels\nfig.add_trace(go.Scatter3d(\n    x=[axis_len*1.1, 0, 0],\n    y=[0, axis_len*1.1, 0],\n    z=[0, 0, axis_len*1.1],\n    mode=\'text\',\n    text=[\'X\', \'Y\', \'Z\'],\n    textfont=dict(size=18, color=[\'red\',\'green\',\'blue\']),\n    hoverinfo=\'none\',\n    showlegend=False\n))\n\n\n\n\n\n# --- Highlight Max-P and Min-P columns (keep existing hover elsewhere) ---\n\n# find rows for max/min P\ni_max = merged_df_c[\'UIF\'].idxmax()\ni_min = merged_df_c[\'UIF\'].idxmin()\nrow_max = merged_df_c.loc[i_max]\nrow_min = merged_df_c.loc[i_min]\n\ndef col_segment_from_row(row):\n    """Get the column end points. Try actual ends from C; else vertical at (xt,yt)."""\n    try:\n        mask = (C[label_col] == row[label_col])\n        if \'Storey\' in C.columns and \'Storey\' in row:\n            mask = mask & (C[\'Storey\'] == row[\'Storey\'])\n        cnd = C[mask].iloc[0]\n        pI, pJ = cnd[\'__I__\'], cnd[\'__J__\']\n        if (pI in coord_lu) and (pJ in coord_lu):\n            (x1, y1, z1) = coord_lu[pI]\n            (x2, y2, z2) = coord_lu[pJ]\n            return [x1, x2], [y1, y2], [z1, z2]\n    except Exception:\n        pass\n    # fallback: vertical between bottom/top z at centroid\n    x = [row[xt_col], row[xt_col]]\n    y = [row[yt_col], row[yt_col]]\n    z = [row[zb_col], row[zt_col]]\n    return x, y, z\n\n\n\ndef highlight_and_label(row, color, text_prefix):\n    x, y, z = col_segment_from_row(row)\n\n    # thick overlay line (no extra hover)\n    fig.add_trace(go.Scatter3d(\n        x=x, y=y, z=z,\n        mode=\'lines\',\n        line=dict(width=5, color=color),\n        hoverinfo=\'skip\',\n        showlegend=False\n    ))\n    \n      \n\n\n\n\n    # text label near the top (no hover)\n    # small offset so it doesn\'t sit inside the line\n    dx = (max(x) - min(x)) * 0.02 or 0.02\n    dy = (max(y) - min(y)) * 0.02 or 0.02\n    dz = (max(z) - min(z)) * 0.02 or 0.02\n    fig.add_trace(go.Scatter3d(\n        x=[x[1] + dx], y=[y[1] + dy], z=[z[1] + dz],\n        mode=\'text\',\n        text=[f"{text_prefix} UIF = {row[\'UIF\']:.3g}"],\n        textfont=dict(size=16, color=color, family="Arial Bold"),\n        hoverinfo=\'skip\',\n        showlegend=False\n    ))\n\n# apply to max and min\nhighlight_and_label(row_max, \'crimson\', \'Max\')\nhighlight_and_label(row_min, \'magenta\',  \'Min\')\n\n\n\n\nfig.update_layout(scene=dict(aspectmode=\'data\'))\nfig.show()\n')


# In[30]:


# plotly_scaled_slab.py - minimalist version (no grid, no axis labels, only hovers)

import pandas as pd
import numpy as np
import plotly.graph_objects as go
from shapely.geometry import Polygon, MultiPolygon, box, Point
from shapely.ops import unary_union
from shapely import affinity as shp_affinity

# (All helper functions and geometry computations unchanged from previous version)
# Assume A, B, C, D, merged_df_c are already prepared in memory.

# --- Only plotting section below for brevity ---

fig = go.Figure()

# Plotly interactive replacement
import numpy as np
import plotly.graph_objects as go

# re-use your helpers: build_coord_lookup, parse_unique_pts, coords_for_ids
coord_lu = build_coord_lookup(A)
for df in (B, C):
    df['__I__'] = df['UniquePtI'].apply(_to_key)
    df['__J__'] = df['UniquePtJ'].apply(_to_key)
D = D.copy()
D['__ids__'] = D['UniquePts'].apply(parse_unique_pts)

# Build line traces for beams+columns (single trace each for performance)
def build_line_trace(df, coord_lu, name):
    xs, ys, zs = [], [], []
    for _, row in df.iterrows():
        pI, pJ = row['__I__'], row['__J__']
        if pI in coord_lu and pJ in coord_lu:
            x1,y1,z1 = coord_lu[pI]; x2,y2,z2 = coord_lu[pJ]
            xs += [x1, x2, None]   # None creates breaks between segments
            ys += [y1, y2, None]
            zs += [z1, z2, None]
    return go.Scatter3d(x=xs, y=ys, z=zs, mode='lines', name=name, line=dict(width=6)
                       )

beam_trace = build_line_trace(B, coord_lu, 'Beams')
beam_trace.update(hoverinfo='skip')

col_trace  = build_line_trace(C, coord_lu, 'Columns')
col_trace['line']['color'] = 'black'
col_trace.update(hoverinfo='skip')

# Nodes (optional)
nx, ny, nz, nlabels = [], [], [], []
for k,(x,y,z) in coord_lu.items():
    nx.append(x); ny.append(y); nz.append(z); nlabels.append(str(k))
nodes = go.Scatter3d(x=nx, y=ny, z=nz, mode='markers', name='Nodes',
                     marker=dict(size=0,opacity=0), text=nlabels, hoverinfo='skip',visible=False)

# Slabs: convert each polygon to triangles by fan triangulation (good for convex/planar)
mesh_traces = []
for _, row in D.iterrows():
    ids = row['__ids__']; verts = coords_for_ids(ids, coord_lu)
    if len(verts) < 3: 
        continue
    verts = np.array(verts)
    x, y, z = verts[:,0], verts[:,1], verts[:,2]

    # Fan triangulation: triangles (0,i,i+1) for i=1..n-2
    i_idxs = []; j_idxs = []; k_idxs = []
    for i in range(1, len(verts)-1):
        i_idxs.append(0); j_idxs.append(i); k_idxs.append(i+1)

    mesh = go.Mesh3d(x=x, y=y, z=z,
                     i=i_idxs, j=j_idxs, k=k_idxs,
                     opacity=0.6, name=f"Slab {row.get('FloorBay','')}",
                     hoverinfo='skip',flatshading=True)
    mesh.update(lighting=dict(ambient=1, diffuse=0, specular=0))
    mesh_traces.append(mesh)
   
    
for mesh in mesh_traces:
    mesh.color = 'lightgreen'   # slab color
    

fig = go.Figure(data=[beam_trace, col_trace, nodes] + mesh_traces)
fig.update_layout(scene=dict(
    xaxis=dict(showgrid=False, showticklabels=False),  # Remove gridlines and ticks from X axis
    yaxis=dict(showgrid=False, showticklabels=False),  # Remove gridlines and ticks from Y axis
    zaxis=dict(showgrid=False, showticklabels=False),  # Remove gridlines and ticks from Z axis
    aspectmode='data'  # Preserve aspect ratio
), title='')



# Hover markers at mid-height
hover_x, hover_y, hover_z, hover_text = [], [], [], []
for _, row in merged_df_c.iterrows():
    hx = row[xt_col]; hy = row[yt_col]
    hz = (row[zt_col] + row[zb_col]) / 2
    txt = f"Storey: {row['Storey']}<br>Column: {row[label_col]}<br>UIF: {row['UIF']:.2f}"
    hover_x.append(hx); hover_y.append(hy); hover_z.append(hz); hover_text.append(txt)

fig.add_trace(go.Scatter3d(x=hover_x, y=hover_y, z=hover_z,
                           mode='markers', marker=dict(size=4, color='white'),
                           hovertext=hover_text, hoverinfo='text', name=''))

# Minimal layout: no grid, no axes, no labels
fig.update_layout(
    scene=dict(
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        zaxis=dict(visible=False),
        bgcolor='white'
    ),
    paper_bgcolor='white',
    plot_bgcolor='white',
    showlegend=False,
    margin=dict(l=0, r=0, t=0, b=0)
)


# Set reasonable camera (optional)
fig.update_layout(scene_camera=dict(eye=dict(x=1.25, y=1.25, z=0.8)))


# ---------- Add XYZ axes with arrowheads ----------
axis_len = 1  # adjust based on your model scale
arrow_size = 0.2 * axis_len  # arrowhead size

# Axis lines
fig.add_trace(go.Scatter3d(
    x=[0, axis_len], y=[0, 0], z=[0, 0],
    mode='lines',
    line=dict(color='red', width=10),
    hoverinfo='none',
    showlegend=False
))
fig.add_trace(go.Scatter3d(
    x=[0, 0], y=[0, axis_len], z=[0, 0],
    mode='lines',
    line=dict(color='green', width=10),
    hoverinfo='none',
    showlegend=False
))
fig.add_trace(go.Scatter3d(
    x=[0, 0], y=[0, 0], z=[0, axis_len],
    mode='lines',
    line=dict(color='blue', width=10),
    hoverinfo='none',
    showlegend=False
))

# Arrowheads using cones
fig.add_trace(go.Cone(
    x=[axis_len], y=[0], z=[0],
    u=[1], v=[0], w=[0],
    sizemode="absolute",
    sizeref=arrow_size,
    colorscale=[[0, 'red'], [1, 'red']],
    showscale=False,
    anchor="tail",
    hoverinfo="skip"
))
fig.add_trace(go.Cone(
    x=[0], y=[axis_len], z=[0],
    u=[0], v=[1], w=[0],
    sizemode="absolute",
    sizeref=arrow_size,
    colorscale=[[0, 'green'], [1, 'green']],
    showscale=False,
    anchor="tail",
    hoverinfo="skip"
))
fig.add_trace(go.Cone(
    x=[0], y=[0], z=[axis_len],
    u=[0], v=[0], w=[1],
    sizemode="absolute",
    sizeref=arrow_size,
    colorscale=[[0, 'blue'], [1, 'blue']],
    showscale=False,
    anchor="tail",
    hoverinfo="skip"
))

# Axis labels
fig.add_trace(go.Scatter3d(
    x=[axis_len*1.1, 0, 0],
    y=[0, axis_len*1.1, 0],
    z=[0, 0, axis_len*1.1],
    mode='text',
    text=['X', 'Y', 'Z'],
    textfont=dict(size=18, color=['red','green','blue']),
    hoverinfo='none',
    showlegend=False
))





# --- Highlight Max-P and Min-P columns (keep existing hover elsewhere) ---

# find rows for max/min P
i_max = merged_df_c['UIF'].idxmax()
i_min = merged_df_c['UIF'].idxmin()
row_max = merged_df_c.loc[i_max]
row_min = merged_df_c.loc[i_min]


max_UIF   = row_max['UIF']
max_col = row_max['ColumnBay']

min_UIF   = row_min['UIF']
min_col = row_min['ColumnBay']

fig.add_annotation(
    text=f"<b>Maximum and Minimum</b><br>"
         #f"Max UIF: {max_UIF:.2f} (Col {max_col})<br>"
         #f"Min UIF: {min_UIF:.2f} (Col {min_col})",
         f"Max UIF: {max_UIF:.2f} <br>"
         f"Min UIF: {min_UIF:.2f}",
    xref="paper", yref="paper",
    x=0.9, y=0.9,
    showarrow=False,
    bordercolor="black",
    borderwidth=1,
    bgcolor="white",
)




def col_segment_from_row(row):
    """Get the column end points. Try actual ends from C; else vertical at (xt,yt)."""
    import re
    import pandas as pd

    try:
        # 1) Match by column label first
        mask = (C[label_col] == row[label_col])

        # 2) Normalize storey on the ROW side -> numeric
        row_storey = None
        if 'Storey' in row and pd.notna(row['Storey']):
            try:
                row_storey = int(row['Storey'])
            except Exception:
                pass
        if row_storey is None and ('Story' in row) and pd.notna(row['Story']):
            m = re.search(r'\d+', str(row['Story']))
            if m:
                row_storey = int(m.group())

        # 3) Normalize storey on the C side -> numeric Series
        C_storey_numeric = None
        if 'Storey' in C.columns:
            C_storey_numeric = pd.to_numeric(C['Storey'], errors='coerce')
        elif 'Story' in C.columns:
            C_storey_numeric = (
                C['Story'].astype(str).str.extract(r'(\d+)')[0].astype(float)
            )

        # 4) If both sides have storey info, include it in the mask
        if (row_storey is not None) and (C_storey_numeric is not None):
            mask = mask & (C_storey_numeric == float(row_storey))

        # 5) Pull the first matching ETABS segment
        cnd = C[mask].iloc[0]
        pI, pJ = cnd['__I__'], cnd['__J__']
        if (pI in coord_lu) and (pJ in coord_lu):
            (x1, y1, z1) = coord_lu[pI]
            (x2, y2, z2) = coord_lu[pJ]
            return [x1, x2], [y1, y2], [z1, z2]

    except Exception:
        pass

    # Fallback: vertical between bottom/top z at centroid
    x = [row[xt_col], row[xt_col]]
    y = [row[yt_col], row[yt_col]]
    z = [row[zb_col], row[zt_col]]
    return x, y, z



def highlight_and_label(row, color, text_prefix):
    x, y, z = col_segment_from_row(row)

    # thick overlay line (no extra hover)
    fig.add_trace(go.Scatter3d(
        x=x, y=y, z=z,
        mode='lines',
        line=dict(width=5, color=color),
        hoverinfo='skip',
        showlegend=False
    ))
    
      

    # text label near the top (no hover)
    # small offset so it doesn't sit inside the line
    dx = (max(x) - min(x)) * 0.02 or 0.02
    dy = (max(y) - min(y)) * 0.02 or 0.02
    dz = (max(z) - min(z)) * 0.02 or 0.02
    fig.add_trace(go.Scatter3d(
        x=[x[1] + dx], y=[y[1] + dy], z=[z[1] + dz],
        mode='text',
        text=[f"{text_prefix} UIF = {row['UIF']:.4g}"],
        textfont=dict(size=16, color=color, family="Arial Bold"),
        hoverinfo='skip',
        showlegend=False
    ))

# apply to max and min
highlight_and_label(row_max, 'crimson', 'Max')
highlight_and_label(row_min, 'magenta',  'Min')


#print(CF.to_string(index=False, float_format=lambda v: f"{v:.4f}"))
CF.to_excel("⬇ Download Axial Force, Tributary Area and UIF data.xlsx", index=False)
from IPython.display import FileLink, display
display(FileLink("Axial Force, Tributary Area and UIF data.xlsx"))  # clickable download link

fig.update_layout(scene=dict(aspectmode='data'))
fig.show()


# In[31]:


# ---- Minimal GUI wrapper for your last plotting cell (grey background) ----
import ipywidgets as w
from IPython.display import display, HTML

# one-time CSS for grey panel
display(HTML("""
<style>
.rc-panel { background:#e6e6e6; border:1px solid #ccc; border-radius:8px; padding:12px; }
.rc-title { text-align:center; color:#1f5fbf; font-weight:800; line-height:1.25; margin:0 0 10px 0; }
.rc-title .l1 { font-size:26px; }
.rc-title .l2 { font-size:22px; }
.rc-title .l3 { font-size:30px; }
.rc-title .l0 { font-size:14px; }
</style>
"""))

display(HTML("""
<style>
.custom-run-btn {
    font-size:18px !important;
    font-weight:700 !important;
}
</style>
"""))


# widgets
run_btn   = w.Button(description="Run Analysis", layout=w.Layout(width='180px', height='40px', border='1.5px solid black'), 
                     style={'font_weight': 'bold'})
status_ht = w.HTML(value="")
fig_out   = w.Output()   # where the Plotly figure will appear
#fig_out   = w.Output(layout=w.Layout(height='200px'))   # or 600px, 900px as per your screen
dl_out    = w.Output()   # where the download link will appear

# header
header = w.HTML("""
<div class="rc-title">
  <div class="l1">Unified Identification Factor (UIF)</div>
  <div class="l2">Interactive 3D Frame</div>
  <div class="l2">with Max/Min UIF highlights</div>
</div>
""")

# container
panel = w.VBox(
    [header,
     w.HBox([run_btn, status_ht], layout=w.Layout(align_items='center', gap='12px')),
     fig_out,
     dl_out],
    layout=w.Layout(width='100%')
)
panel.add_class("rc-panel")
display(panel)





# ---------- your plotting routine wrapped in a function ----------
def render_plot_and_link():
    """
    Executes your existing plotting code and returns nothing.
    Renders figure + download link into fig_out and dl_out.
    Assumes A,B,C,D,merged_df_c and helper functions/column names already exist in memory.
    """
    import pandas as pd
    import numpy as np
    import plotly.graph_objects as go
    from IPython.display import FileLink, display as _display

    # clear previous outputs
    fig_out.clear_output()
    dl_out.clear_output()

    # --------- BEGIN: your original plotting section (unchanged except scoped to outputs) ---------
    # (All helper functions and geometry computations unchanged from previous version)
    # Assume A, B, C, D, merged_df_c are already prepared in memory.

    # Plotly interactive replacement
    coord_lu = build_coord_lookup(A)
    for df in (B, C):
        df['__I__'] = df['UniquePtI'].apply(_to_key)
        df['__J__'] = df['UniquePtJ'].apply(_to_key)
    _D = D.copy()
    _D['__ids__'] = _D['UniquePts'].apply(parse_unique_pts)

    # Build line traces for beams+columns (single trace each for performance)
    def build_line_trace(df, coord_lu, name):
        xs, ys, zs = [], [], []
        for _, row in df.iterrows():
            pI, pJ = row['__I__'], row['__J__']
            if pI in coord_lu and pJ in coord_lu:
                x1,y1,z1 = coord_lu[pI]; x2,y2,z2 = coord_lu[pJ]
                xs += [x1, x2, None]   # None creates breaks between segments
                ys += [y1, y2, None]
                zs += [z1, z2, None]
        return go.Scatter3d(x=xs, y=ys, z=zs, mode='lines', name=name, line=dict(width=6))

    beam_trace = build_line_trace(B, coord_lu, 'Beams')
    beam_trace.update(hoverinfo='skip')

    col_trace  = build_line_trace(C, coord_lu, 'Columns')
    col_trace['line']['color'] = 'black'
    col_trace.update(hoverinfo='skip')

    # Nodes (optional)
    nx, ny, nz, nlabels = [], [], [], []
    for k,(x,y,z) in coord_lu.items():
        nx.append(x); ny.append(y); nz.append(z); nlabels.append(str(k))
    nodes = go.Scatter3d(x=nx, y=ny, z=nz, mode='markers', name='Nodes',
                         marker=dict(size=0), text=nlabels, hoverinfo='skip',visible=False)

    # Slabs: fan triangulation
    mesh_traces = []
    for _, row in _D.iterrows():
        ids = row['__ids__']; verts = coords_for_ids(ids, coord_lu)
        if len(verts) < 3: 
            continue
        verts = np.array(verts)
        x, y, z = verts[:,0], verts[:,1], verts[:,2]

        i_idxs = []; j_idxs = []; k_idxs = []
        for i in range(1, len(verts)-1):
            i_idxs.append(0); j_idxs.append(i); k_idxs.append(i+1)

        mesh = go.Mesh3d(x=x, y=y, z=z,
                         i=i_idxs, j=j_idxs, k=k_idxs,
                         opacity=0.9, name=f"Slab {row.get('FloorBay')}",
                         hoverinfo='skip')
        mesh.color = 'lightgreen'
        mesh_traces.append(mesh)

    fig = go.Figure(data=[beam_trace, col_trace, nodes] + mesh_traces)
    fig.update_layout(scene=dict(
        xaxis=dict(showgrid=False, showticklabels=False),
        yaxis=dict(showgrid=False, showticklabels=False),
        zaxis=dict(showgrid=False, showticklabels=False),
        aspectmode='data'
    ), title='');
    fig.update_layout(autosize=True)
    _ = 0  # suppress echo

    # Hover markers at mid-height
    hover_x, hover_y, hover_z, hover_text = [], [], [], []
    for _, row in merged_df_c.iterrows():
        hx = row[xt_col]; hy = row[yt_col]
        hz = (row[zt_col] + row[zb_col]) / 2
        txt = f"Storey: {row['Storey']}<br>Column: {row[label_col]}<br>UIF: {row['UIF']:.2f}"
        hover_x.append(hx); hover_y.append(hy); hover_z.append(hz); hover_text.append(txt)

    fig.add_trace(go.Scatter3d(x=hover_x, y=hover_y, z=hover_z,
                               mode='markers', marker=dict(size=4, color='white'),
                               hovertext=hover_text, hoverinfo='text', name=''))

    # minimal layout
    fig.update_layout(
        scene=dict(xaxis=dict(visible=False), yaxis=dict(visible=False), zaxis=dict(visible=False), bgcolor='#e6e6e6'),
        paper_bgcolor='#e6e6e6', plot_bgcolor='#e6e6e6', showlegend=False, margin=dict(l=0, r=0, t=0, b=0)
    );
    _ = 0  # suppress echo
    fig.update_layout(scene_camera=dict(eye=dict(x=1.25, y=1.25, z=0.8)));
    _ = 0  # suppress echo

    # XYZ axes lines + cones + labels
    axis_len   = 1
    arrow_size = 0.2 * axis_len
    fig.add_trace(go.Scatter3d(x=[0, axis_len], y=[0, 0], z=[0, 0], mode='lines',
                               line=dict(color='red', width=10), hoverinfo='none', showlegend=False))
    fig.add_trace(go.Scatter3d(x=[0, 0], y=[0, axis_len], z=[0, 0], mode='lines',
                               line=dict(color='green', width=10), hoverinfo='none', showlegend=False))
    fig.add_trace(go.Scatter3d(x=[0, 0], y=[0, 0], z=[0, axis_len], mode='lines',
                               line=dict(color='blue', width=10), hoverinfo='none', showlegend=False))

    fig.add_trace(go.Cone(x=[axis_len], y=[0], z=[0], u=[1], v=[0], w=[0],
                          sizemode="absolute", sizeref=arrow_size,
                          colorscale=[[0, 'red'], [1, 'red']], showscale=False, anchor="tail", hoverinfo="skip"))
    fig.add_trace(go.Cone(x=[0], y=[axis_len], z=[0], u=[0], v=[1], w=[0],
                          sizemode="absolute", sizeref=arrow_size,
                          colorscale=[[0, 'green'], [1, 'green']], showscale=False, anchor="tail", hoverinfo="skip"))
    fig.add_trace(go.Cone(x=[0], y=[0], z=[axis_len], u=[0], v=[0], w=[1],
                          sizemode="absolute", sizeref=arrow_size,
                          colorscale=[[0, 'blue'], [1, 'blue']], showscale=False, anchor="tail", hoverinfo="skip"))
    fig.add_trace(go.Scatter3d(x=[axis_len*1.1, 0, 0], y=[0, axis_len*1.1, 0], z=[0, 0, axis_len*1.1],
                               mode='text', text=['X','Y','Z'],
                               textfont=dict(size=18, color=['red','green','blue']),
                               hoverinfo='none', showlegend=False))

    # --- Highlight Max-P and Min-P columns ---
    i_max = merged_df_c['UIF'].idxmax()
    i_min = merged_df_c['UIF'].idxmin()
    row_max = merged_df_c.loc[i_max]
    row_min = merged_df_c.loc[i_min]

    def col_segment_from_row(row):
        import re, pandas as pd
        try:
            mask = (C[label_col] == row[label_col])
            row_storey = None
            if 'Storey' in row and pd.notna(row['Storey']):
                try: row_storey = int(row['Storey'])
                except: pass
            if row_storey is None and ('Story' in row) and pd.notna(row['Story']):
                m = re.search(r'\d+', str(row['Story']))
                if m: row_storey = int(m.group())
            C_storey_numeric = None
            if 'Storey' in C.columns:
                C_storey_numeric = pd.to_numeric(C['Storey'], errors='coerce')
            elif 'Story' in C.columns:
                C_storey_numeric = C['Story'].astype(str).str.extract(r'(\d+)')[0].astype(float)
            if (row_storey is not None) and (C_storey_numeric is not None):
                mask = mask & (C_storey_numeric == float(row_storey))
            cnd = C[mask].iloc[0]
            pI, pJ = cnd['__I__'], cnd['__J__']
            if (pI in coord_lu) and (pJ in coord_lu):
                (x1,y1,z1) = coord_lu[pI]; (x2,y2,z2) = coord_lu[pJ]
                return [x1, x2], [y1, y2], [z1, z2]
        except Exception:
            pass
        x = [row[xt_col], row[xt_col]]
        y = [row[yt_col], row[yt_col]]
        z = [row[zb_col], row[zt_col]]
        return x, y, z

    def highlight_and_label(row, color, text_prefix):
        x, y, z = col_segment_from_row(row)
        fig.add_trace(go.Scatter3d(x=x, y=y, z=z, mode='lines',
                                   line=dict(width=5, color=color),
                                   hoverinfo='skip', showlegend=False))
        dx = (max(x)-min(x))*0.02 or 0.02
        dy = (max(y)-min(y))*0.02 or 0.02
        dz = (max(z)-min(z))*0.02 or 0.02
        fig.add_trace(go.Scatter3d(x=[x[1]+dx], y=[y[1]+dy], z=[z[1]+dz],
                                   mode='text',
                                   text=[f"{text_prefix} UIF = {row['UIF']:.4g}"],
                                   textfont=dict(size=18, color=color, family="Arial Bold"),
                                   hoverinfo='skip', showlegend=False))

    highlight_and_label(row_max, 'crimson', 'Max')
    highlight_and_label(row_min, 'magenta',  'Min')

    # annotation box (only P values)
    fig.add_annotation(
        text=(f"<b></b><br>"
              f"Max UIF: {row_max['UIF']:.2f}<br>"
              f"Min UIF: {row_min['UIF']:.2f}"),
        xref="paper", yref="paper", x=0.8, y=1,
        showarrow=False, bordercolor="black", borderwidth=1.5, bgcolor="white",
        font=dict(size=20, color="black", family="Arial Bold")
    )

    fig.update_layout(scene=dict(aspectmode='data'));
    _ = 0  # suppress echo

    # Render figure into output area
    #with fig_out:
        #fig.show()
    # Render figure into grey panel
    with fig_out:
        display(HTML("<div style='background:#e6e6e6; padding:10px; border-radius:8px;'>"))
        #fig.show()
        fig.show(config={'displayModeBar': False})
        display(HTML("</div>"))


    # Save and show download link into its own area
    with dl_out:
        fname = "Axial force, Tributary area and UIF data.xlsx"
        CF.to_excel(fname, index=False)
       # _display(FileLink(fname))
        
        display(HTML(f"""
    <a href="{fname}" download
       style="
           font-size:15px;
           font-weight:bold;
           color:#000000;
           text-decoration:none;
       ">
       ⬇ Download Axial Force, Tributary Area and UIF data
    </a>
    """))
        
    # --------- END: your plotting section ---------

def _on_run(_):
    status_ht.value = "<i>Running…</i>"
    try:
        render_plot_and_link()
        status_ht.value = "<b>Analysis completed</b>"
    except Exception as e:
        fig_out.clear_output(); dl_out.clear_output()
        status_ht.value = f"<span style='color:#b00020;'><b>Error:</b> {e}</span>"

run_btn.on_click(_on_run)


# In[ ]:




