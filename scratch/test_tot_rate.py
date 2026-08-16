import sys

cam_linh_vol_ca1 = 763
cam_linh_vol_ton = 1122
cam_linh_gtc_ca1 = 471
cam_linh_gtc_ton = 414

tot_vol_cam_linh = cam_linh_vol_ca1 + cam_linh_vol_ton
tot_gtc_cam_linh = cam_linh_gtc_ca1 + cam_linh_gtc_ton
rate_cam_linh = tot_gtc_cam_linh / tot_vol_cam_linh * 100

print(f"Cam Linh: {tot_gtc_cam_linh} / {tot_vol_cam_linh} = {rate_cam_linh:.2f}%")

di_linh_vol_ca1 = 603
di_linh_vol_ton = 556
di_linh_gtc_ca1 = 359
di_linh_gtc_ton = 200

tot_vol_di_linh = di_linh_vol_ca1 + di_linh_vol_ton
tot_gtc_di_linh = di_linh_gtc_ca1 + di_linh_gtc_ton
rate_di_linh = tot_gtc_di_linh / tot_vol_di_linh * 100

print(f"Di Linh: {tot_gtc_di_linh} / {tot_vol_di_linh} = {rate_di_linh:.2f}%")

