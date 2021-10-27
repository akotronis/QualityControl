# SKU Quality Control Application

The purpose of SKU Quality Control application is to investigate the atypical purchases that are
recorded in one audit period.
The process of finding atypical purchases reaches the SKU level per store
of the _Emrc Retail Audit sample._

The distance measure which is used to implement the above, based on Shanon’s theory, is the
following:

<p style="text-align: center;"><img src="https://latex.codecogs.com/gif.latex?\text{             }D(P_{t+1},P_t)=P_{t+1}\cdot\ln(P_{t+1}/P_t)+P_t-P_{t+1}\text{             }" /></p>
$$D(P_{t+1},P_t)=P_{t+1}\cdot\ln(P_{t+1}/P_t)+P_t-P_{t+1}.$$

## Gui
