# ribbon-twenty-oh-seven
## Key Observations
Icons are all available by downloading and extracting from the ZIP file. All icons are extacted at 16px by 16px, 32px by 32px resolutions. Icons are natively only ONE of their two available resolutions. First look through the `* 32.png` filtered files, if the icon you want looks blurry in the 32 variant, it will look crisp and correct with the same file name but ending in ` 16.png`. The icons also contain the same image with different names, so duplication is present in the data set.

## Developer Comments
The master list of icon names in the repo results in duplication when used with the VBA macro script. The VBA script needs to be attached to an Excel workbook with the master list of icons present as column A of the active sheet. The macro should be edited for pulling at size 16 or pulling at size 32. There is no way from the APIs to determine or query the correct icon size. The macro can run in different model-years of Excel for different results, macros must be enabled to generate the icons yourself, they will be placed in `/Temp`.
