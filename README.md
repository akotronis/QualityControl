# SKU Quality Control Application

## Introduction

The purpose of _SKU Quality Control application_ is to investigate the atypical purchases that are
recorded in one audit period.
The process of finding atypical purchases reaches the SKU level per store
of the _Emrc Retail Audit sample._

The distance measure which is used to implement the above, based on Shanon’s theory, is the
following:

<img src="https://latex.codecogs.com/gif.latex?%5Ctext%7B%20%7DD%28P_%7Bt&plus;1%7D%2CP_t%29%3DP_%7Bt&plus;1%7D%5Ccdot%5Cln%28P_%7Bt&plus;1%7D/P_t%29&plus;P_t-P_%7Bt&plus;1%7D%5Ctext%7B%20%7D" />

## An example

A typical example that describes this property is the following:
We have 2 stores where the purchases for the same SKU in 2 consecutive periods are

<img src="https://latex.codecogs.com/gif.latex?%5Ctext%7Bstore%20A%3A%20%7DP_t%3D1%2C%20P_%7Bt&plus;1%7D%3D2%5Ctext%7B%2C%20and%20store%20B%3A%20%7DP_t%3D5%2C%20P_%7Bt&plus;1%7D%3D10." />

We immediately observe the doubling of the purchases in the period t + 1, that is, we have a 100%
increase in both stores. Here, the increase in store A should not worry us because it is reasonable and
expected to buy an additional SKU and we must pay attention to store B, something that the classic
distance measures do not indicate.
However, by using the above type of distance, we obtain

<img src="https://latex.codecogs.com/gif.latex?D%28P_%7Bt&plus;1%7D%2CP_t%29%3D0.386%5Ctext%7B%20for%20the%20store%20A%20and%20%7DD%28P_%7Bt&plus;1%7D%2CP_t%29%3D1.931%5Ctext%7B%20for%20the%20store%20B%2C%7D" />

## Implementation

This application implements the following 2 procedures.

1. The application builds for each SKU, per audit period (monthly, bi-monthly, etc.) and per cluster,
   the distribution of the distance measure <img src="https://latex.codecogs.com/gif.latex?D%28P_%7Bt&plus;1%7D%2C%20P_t%29" />. Each such distribution is built with the past
   data and updated every audit period with the new data. From the study of these distributions the
   percentiles 90%, 95% and 99% are calculated.
2. For all stores "i" of the audit period, the corresponding <img src="https://latex.codecogs.com/gif.latex?D_i%28P_%7Bt&plus;1%7D%2C%20P_t%29" /> is calculated per SKU and
   compared with the 3 critical percentiles of the corresponding <img src="https://latex.codecogs.com/gif.latex?D%28P_%7Bt&plus;1%7D%2C%20P_t%29" />.
   If this Di is greater
   than one of the 3 percentiles then it is characterized as atypical and the application suggests 2
   optimum alternatives.

## Gui

The application performs analysis on user input files and results are stored in an **sqlite3** database. The user interacts with the application through a GUI and he can import files for analysis, export files, delete database contents and read the application documentation. Application info is displayed in the console.

### Main Interface

Below is the application main interface. The user can navigate from the top menu and select an
action.

![](resources/01_main_interface.jpg)

#### Console

The various messages produced by user actions are displayed on the console.

- The console can be cleared by selecting Console → Clear from the top level dropdown menu or by clicking on the console, pressing Ctrl+A and then Delete.
- The contents of the console can be copied by clicking on the console, pressing Ctrl+A, then Ctrl+C and then Ctrl+V to paste them anywhere you like.

#### Actions

The application expects actions in the following logical order:

1. Import Clusters
2. Import SKUS for analysis
3. Import Outlets to perform analysis based on the previous step

The application creates a database file (**db.sqlite3**) in the same folder of the executable, where the
imported items are stored. If this file is deleted, it will be automatically created again the next time the
application is launched and will have to be populated again with entries from new imports.

#### About

- Select _About → Database Current Status_ to print counts for the database tables on the console.
- Select _About → Importing_ to print information about importing files (see below).
- Select _About → SKU Quality Control Application_ to print quick info about the app.
- Select _About → Read the Docs_ to launch the documentation file.

![](resources/00_About.jpg)

## Importing files

The application accepts as input **.csv** or **.xlsx/.xls** files. The .csv files are imported much faster. In
the case of .xlsx/.xls imports, the application will try to convert the file to .csv before parsing it.

### Clusters

The imported file must follow the rules the user can see by selecting About → Importing → Clusters from the top level dropdown menu.
![](resources/02_about_importing_clusters.jpg)

Select _Import → Clusters_ from the top level dropdown and choose a cluster file to import.
![](resources/03_importing_clusters.jpg)

The appropriate messages will be displayed on the console in the case of success or failure accord-
ingly.
![](resources/04_importing_clusters_success.jpg)
![](resources/05_importing_clusters_failure.jpg)

Note that:

1. The imported file must comply with the rules mentioned above,
2. Every time a new file is submitted, **all existing clusters will be deleted and replaced by the new ones**.

### SKUs

The imported file must follow the rules the user can see by selecting _About → Importing → SKUs_ from the top level dropdown menu
![](resources/06_about_importing_skus.jpg)

Select _Import → SKUs →_ and the period type for the file you want to import from the top level
dropdown.
![](resources/07_importing_skus_1.jpg)

A popup will appear with the period type that was selected from the menu. Make sure to select the file of the correct period type, or an error message will appear while parsing it.
![](resources/07_importing_skus_2.jpg)

The appropriate success/error messages will appear while parsing the file. If parsing is successful,
a progress window will appear displaying the progress of the SKU analysis. Wait until the process is
finished.
![](resources/07_importing_skus_3.jpg)
