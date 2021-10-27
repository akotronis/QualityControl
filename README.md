# SKU Quality Control Application

## Introduction

The purpose of SKU Quality Control application is to investigate the atypical purchases that are
recorded in one audit period.
The process of finding atypical purchases reaches the SKU level per store
of the _Emrc Retail Audit sample._

The distance measure which is used to implement the above, based on Shanon’s theory, is the
following:

<img src="https://latex.codecogs.com/gif.latex?%5Ctext%7B%20%7DD%28P_%7Bt&plus;1%7D%2CP_t%29%3DP_%7Bt&plus;1%7D%5Ccdot%5Cln%28P_%7Bt&plus;1%7D/P_t%29&plus;P_t-P_%7Bt&plus;1%7D%5Ctext%7B%20%7D" />

## An example

A typical example that describes this property is the following:
We have 2 stores where the purchases for the same SKU in 2 consecutive periods are

<img src="https://latex.codecogs.com/gif.latex?\text{store A: }P_t=1, P_{t+1}=2\text{, and store B: }P_t=5, P_{t+1}=10." />
## Gui
