# lighting_supply_assembly_sorting

This project was commissioned by a company in the lighting supplies industry.  The company was in trasition from their European distribution chain to their North American distribution chain, during which the North American distributor needed to consolidate all product codes and names to fit their inventory system.  The task involves a product price list of over 15,000 products, and a backend assembly parts list of over 300,000 components. 

Due to the amount of entries involved, manual adjustment was not an option.  There were also different naming schemes that applies to different product families.  Eg. colour temperature in the naming roughly follows the follow criteria:
-   GROUP 1 uses 2700K, 3000k, 4000K to represent color temperature. it translates to 2, 3, and 4 in product code
-   ANDRE use rule of HW, WW, NW, CW; LAST DIGIT
-   ARIK2 & ARIK6 use rule of HW, WW, NW, CW; Third LAST DIGIT (2 letters after it)
-   BANG use rule of HW, WW, NW, CW; LAST DIGIT
-   ELLE (Code starts with BLP & BLS) is an exception. The description is using 2700K, 3000K, 4200K, and 6000K as description, NOT HW, WW, NW, and CW. But it’s using code in 0, 1, 2, 3, corresponding to the description, which usually use letter in description, not numbers. In ELLE series, the color temperature is the LAST DIGITAL even though there is one more description after the color temperature. The original code for ELLE is 2700K = 0, 3000K =1, 4200K = 2, 6000K = 3. Now we’re changing to HW = 27, WW = 30, NW = 42, CW = 60.
-   BABAR2 & BABAR6 (code starts with BS) is another exception. Following same instruction as ELLE. The color temperature code is last digit
-   BABA (Code starts with BS also) follow the rule of using HW, WW, NW, CW, LAST DIGIT
-   BABUS went back to using 2700, 3000, 4200, 6000. AND the color temperature is the SECOND LAST DIGIT, NOT LAST
-   BABA (Code starts with BZ) uses the rule of 2700, 3000, 4200, 6000 to descript color temperature (different from BABA with code of BS)
-   CRIS (code starts with CR) uses the rule of 2700, 3000, 4200, 6000 to descript color temperature and it’s on the SECOND LAST DIGIT
-   DIDO (code starts with DD), DAGO2, DAGO6,use the rule of HW, WW, NW, CW, LAST DIGIT
-   FATBOY uses rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   KING (KG) use the rule of HW, WW, NW, CW, LAST DIGIT
-   KOINE2 & KOINE6 use rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   LOTUS2 & LOTUS6 use rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   MARUPE4 use rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   MARUPE12 use rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   SOREL use rule of HW, WW, NW, CW; LAST DIGIT
-   ADAN use rule of HW, WW, NW, CW; LAST DIGIT
-   KOL use rule of HW, WW, NW, CW; LAST DIGIT
-   CADMO6 & CADMO2 use rule of HW, WW, NW, CW; Third LAST DIGIT (2 letters after it)
-   BELEM use rule of HW, WW, NW, CW; LAST DIGIT
-   BAY use rule of HW, WW, NW, CW; LAST DIGIT
-   DAHLIA use rule of HW, WW, NW, CW; LAST DIGIT
-   SKY use rule of HW, WW, NW, CW; LAST DIGIT
-   SELENE use rule of HW, WW, NW, CW; LAST DIGIT
-   TYLA use rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   UMA use rule of HW, WW, NW, CW; LAST DIGIT
-   JENA2 & JENA6 use rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   ALOHA 2 & ALOHA6 & ALOHA14 & ALOHA 524 use rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   AIR6 use rule of HW, WW, NW, CW; LAST DIGIT
-   TERA2, TERA6, TERA7, TERA 524 use rule of HW, WW, NW, CW; SECOND LAST DIGIT
-   ZERINO use rule of HW, WW, NW, CW; LAST DIGIT
-   ZOOM use rule of HW, WW, NW, CW; LAST DIGIT
-   ZERO (ZS) use rule of HW, WW, NW, CW; LAST DIGIT
-   ZERUS use rule of 2700, 3000, 4200, 6000; SECOND LAST DIGIT
-   ZERO (ZW) use rule of HW, WW, NW, CW; LAST DIGIT
-   ZERO (ZZ) use rule of HW, WW, NW, CW; LAST DIGIT

My strategy to tackle this problem is the following:
1. Divide the lists into 2 groups, group 1 being transformations that are straight forward, and group 2 being entries that require filters to funnel them into the correct rules.
2. For items in group 2, set up patterns to 
     - check which family and spec type the product name falls under
	 - rules to break up the naming into segments for further processing
	 - renaming rule for each segment 
3. If a family is found, but not type, separate the entry into an "exceptions" list.
4. For each family found, run the exceptions list with patterns that were set up for other families.  
5. If step 4 fails, search for a common pattern within the exceptions list.  The pattern is saved as a custom rule.
6. Rerun the exceptions list with the custom rule.

This project was done using VBA in Excel.  My code was able to recognize all the preset patterns within the two lists, as well as outliers that were not define by the company.  All 315,000+ items were sucessfully processed in the two files.
