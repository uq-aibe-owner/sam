#cd("C:\\Users\\jaber\\OneDrive\\Documents\\AIBE\\sam")
#using Ipopt: optimize!
using Base: Int64
using XLSX: length
#=using Pkg
Pkg.add("Tables")
Pkg.add("DataFrames")
Pkg.add("XLSX")
Pkg.add("ExcelReaders")
Pkg.add("JuMP")
Pkg.add("Ipopt")
=#
using XLSX, ExcelReaders, DataFrames, Tables, JuMP, Ipopt;
IOSource = ExcelReaders.readxlsheet("5209055001DO001_201819.xls", "Table 5");

#filepath cross system compatability code
if Sys.KERNEL === :NT || kern === :Windows
    pathmark = "\\"
else
    pathmark = "/"
end

#indexing vectors for initial data import groups
intermediaryTotalsCol = findall(x -> occursin("T4", x), string.(IOSource[3,:]));
intermediaryTotalsRow = findall(x -> occursin("T1", x), string.(IOSource[:,1]));
finalTotalsCol = findall(x -> occursin("T6", x), string.(IOSource[3,:]));
finalTotalsRow = findall(x -> occursin("Australian Production", x), string.(IOSource[:,2]));
finalDemandCol = findall(x -> occursin('Q', x), string.(IOSource[3,:]));
factorRow = findall(x -> occursin('P', x), string.(IOSource[:,1]));
IOSourceCol = vcat(intermediaryTotalsCol, finalDemandCol, finalTotalsCol);
IOSourceRow = vcat(intermediaryTotalsRow, factorRow[1:length(factorRow)-1], finalTotalsRow);
#initialising IO
IO = zeros(length(IOSourceRow), length(IOSourceCol));
#import numerical data into IO
IO[1:length(IOSourceRow), 1:length(IOSourceCol)] = IOSource[IOSourceRow, IOSourceCol];
#creating vectors of titles for IO
IOCodeRow = IOSource[IOSourceRow, 1];
IOCodeRow[length(IOSourceRow)] = "T2";
IOCodeCol = IOSource[3, IOSourceCol];
IONameRow = IOSource[IOSourceRow, 2];
IONameCol = IOSource[2, IOSourceCol];

#code to sum public and private entities into one collumn
investment = findall(x -> occursin("Capital Formation", x), IONameCol);
IO[:, investment[1]]=sum(eachcol(IO[:, investment[1:2]]));
IO = IO[:,Not(investment[2])];
#alter title vectors accordingly (include Q in total investment collumn in IOcode)
IOCodeCol[investment[1]] = "Q3+Q4";
IOCodeCol = IOCodeCol[Not(investment[2])];
IONameCol[investment[1]] = "Private and Public Gross Fixed Capital Formation";
IONameCol = IONameCol[Not(investment[2])];
#creating a dictionary for the index of each collumn and row in IO by IOCode
IOColDict = Dict(IOCodeCol .=> [1:1:8;]);
IORowDict = Dict(IOCodeRow .=> [1:1:8;]);
IOCapForm = findall(x -> occursin("Capital Formation", x), IONameCol);
IOChangeInv = findall(x -> occursin("Changes in Inventories", x), IONameCol);

#importing relevant ASNA data for table 5
ASNAHouseCap = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204039_Household_Capital_Account.xls", "Data1");
ASNANonFinCap = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204018_NonFin_Corp_Capital_Account.xls", "Data1");
ASNAFinCap = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204026_Fin_Corp_Capital_Account.xls", "Data1");
ASNAGovCap = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204032_GenGov_Capital_Account.xls", "Data1");
ASNAYearRow = findall(x -> occursin("2019", x), string.(ASNAHouseCap[:,1]));

#table 5
#creating table 5a - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection a is fixed capital expenditure
table5aNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5aNameRow = ["Domestic Commodities", "Imported Commodities, complementary", "Imported Commodities, competing", 
"Taxes less subsidies on products", "Other taxes less subsidies on investment", "Total indirect taxes", 
"Total fixed capital expenditure"];
table5aRowDict = Dict(table5aNameRow .=> [1:1:length(table5aNameRow);]);
table5aColDict = Dict(table5aNameCol .=> [1:1:length(table5aNameCol);]);
table5a = zeros(length(table5aNameRow), length(table5aNameCol));

#filling in totals collumn from corresponding IO data
table5a[table5aRowDict["Domestic Commodities"], table5aColDict["Total"]] = sum(IO[IORowDict["T1"],IOCapForm]);
table5a[table5aRowDict["Imported Commodities, complementary"], table5aColDict["Total"]] = sum(IO[IORowDict["P5"],IOCapForm]);
table5a[table5aRowDict["Imported Commodities, competing"], table5aColDict["Total"]] = sum(IO[IORowDict["P6"],IOCapForm]);
table5a[table5aRowDict["Taxes less subsidies on products"], table5aColDict["Total"]] = sum(IO[IORowDict["P3"],IOCapForm]);
table5a[table5aRowDict["Other taxes less subsidies on investment"], table5aColDict["Total"]] = sum(IO[IORowDict["P4"],IOCapForm]);
table5aTaxes = findall(x -> occursin("taxes", lowercase(x)), table5aNameRow);
table5aTaxes = table5aTaxes[Not(3)];
table5a[table5aRowDict["Total indirect taxes"], table5aColDict["Total"]] = sum(table5a[table5aTaxes,table5aColDict["Total"]]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Total"]] = sum(table5a[Not(table5aRowDict["Total indirect taxes"]),table5aColDict["Total"]]);

#creating index variables for the measurements that we want
ASNAHouseCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNAHouseCap[1,:]));
ASNANonFinCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNANonFinCap[1,:]));
ASNAFinCapTotCapForm = findall(x -> occursin("Gross fixed capital formation ;", x), string.(ASNAFinCap[1,:]));
ASNAGenGovCapTotCapForm = findall(x -> occursin("General government ;  Gross fixed capital formation ;", x), string.(ASNAGovCap[1,:]));

#filling in totals row from ASNA Data
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Households"]] = first(ASNAHouseCap[ASNAYearRow, ASNAHouseCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Non-Financial Corporations"]] = first(ASNANonFinCap[ASNAYearRow, ASNANonFinCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["Financial Corporations"]] = first(ASNAFinCap[ASNAYearRow, ASNAFinCapTotCapForm]);
table5a[table5aRowDict["Total fixed capital expenditure"], table5aColDict["General Government"]] = first(ASNAGovCap[ASNAYearRow, ASNAGenGovCapTotCapForm]);

#filling in non-total values note- i think its including the taxes subtotal in the distribution
for i in [1:1:length(table5aRowDict)-1;];
    for ring in [1:1:length(table5aColDict)-1;];
        table5a[i,ring] = table5a[table5aRowDict["Total fixed capital expenditure"],ring]*table5a[i,table5aColDict["Total"]] / table5a[length(table5aRowDict),length(table5aColDict)];
        table5a[table5aRowDict["Total indirect taxes"], ring] = sum(table5a[table5aTaxes, ring]);
    end
end

#creating table 5b - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection b is fixed capital expenditure
table5bNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5bNameRow = ["Domestic Commodities", "Imported Commodities, complementary", "Imported Commodities, competing", 
"Taxes less subsidies on products", "Total change in inventories"];
table5bRowDict = Dict(table5bNameRow .=> [1:1:length(table5bNameRow);]);
table5bColDict = Dict(table5bNameCol .=> [1:1:length(table5bNameCol);]);
table5b = zeros(length(table5bNameRow), length(table5bNameCol));

#filling in totals collumn from corresponding IO data
table5b[table5bRowDict["Domestic Commodities"], table5bColDict["Total"]] = sum(IO[IORowDict["T1"],IOChangeInv]);
table5b[table5bRowDict["Imported Commodities, complementary"], table5bColDict["Total"]] = sum(IO[IORowDict["P5"],IOChangeInv]);
table5b[table5bRowDict["Imported Commodities, competing"], table5bColDict["Total"]] = sum(IO[IORowDict["P6"],IOChangeInv]);
table5b[table5bRowDict["Taxes less subsidies on products"], table5bColDict["Total"]] = sum(IO[IORowDict["P3"],IOChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Total"]] = sum(table5b[:,table5bColDict["Total"]]);

#creating index variables for the measurements that we want
ASNAHouseCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNAHouseCap[1,:]));
ASNANonFinCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNANonFinCap[1,:]));
ASNAFinCapChangeInv = findall(x -> occursin("Changes in inventories ;", x), string.(ASNAFinCap[1,:]));
ASNAGenGovCapChangeInv = findall(x -> occursin("General government ;  Changes in inventories ;", x), string.(ASNAGovCap[1,:]));

#filling in totals row from ASNA Data
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Households"]] = first(ASNAHouseCap[ASNAYearRow, ASNAHouseCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Non-Financial Corporations"]] = first(ASNANonFinCap[ASNAYearRow, ASNANonFinCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["Financial Corporations"]] = first(ASNAFinCap[ASNAYearRow, ASNAFinCapChangeInv]);
table5b[table5bRowDict["Total change in inventories"], table5bColDict["General Government"]] = first(ASNAGovCap[ASNAYearRow, ASNAGenGovCapChangeInv]);

#calculate non-total values with lagrangian optimisation
table5bScalingFact = abs(minimum(table5b)) * 2;
mod5b = Model(Ipopt.Optimizer);
@variable(mod5b, x[1:(length(table5bNameRow)-1), 1:(length(table5bNameCol)-1)]);
@NLobjective(mod5b, Min, sum((x[i,j] - table5bScalingFact) ^ 2 for i in 1:(length(table5bNameRow)-1), j in 1:(length(table5bNameCol)-1)));
for i in 1:(length(table5bNameRow)-1);
    @constraint(mod5b, sum(x[:,i]) == table5b[table5bRowDict["Total change in inventories"],i]+table5bScalingFact);
end;
for i in 1:(length(table5bNameCol)-1);
    @constraint(mod5b, sum(x[i,:]) == table5b[i,table5bColDict["Total"]]+table5bScalingFact);
end;
optimize!(mod5b);

#plug back into table 5b
table5b[1:(length(table5bNameRow)-1),1:(length(table5bNameCol)-1)]=value.(x).-table5bScalingFact/4;


#creating table 5c - allocation of investment expenditure (broken into subsections for dict referencing purposes)
#subsection c is totals
table5cNameCol = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "Total"];
table5cNameRow = ["Domestic Commodities", "Imported Commodities", "Taxes less subsidies on products", "Other taxes less subsidies on investment", "Total investment expenditure"];
table5cRowDict = Dict(table5cNameRow .=> [1:1:length(table5cNameRow);]);
table5cColDict = Dict(table5cNameCol .=> [1:1:length(table5cNameCol);]);
table5c = zeros(length(table5cNameRow), length(table5cNameCol));

#do totals calcuations to get all values in 5c
table5c[table5cRowDict["Domestic Commodities"],:] = (table5a[table5aRowDict["Domestic Commodities"],:] +
table5b[table5bRowDict["Domestic Commodities"],:]);
table5c[table5cRowDict["Imported Commodities"],:] = sum(eachrow(table5a[[table5aRowDict["Imported Commodities, competing"],table5aRowDict["Imported Commodities, complementary"]],:] +
table5b[[table5bRowDict["Imported Commodities, competing"],table5bRowDict["Imported Commodities, complementary"]],:]));
table5c[table5cRowDict["Taxes less subsidies on products"],:] = table5a[table5aRowDict["Taxes less subsidies on products"],:];
table5c[table5cRowDict["Other taxes less subsidies on investment"],:] = (table5a[table5aRowDict["Other taxes less subsidies on investment"],:] +
table5b[table5bRowDict["Taxes less subsidies on products"],:]);
table5c[table5cRowDict["Total investment expenditure"],:] = (table5a[table5aRowDict["Total fixed capital expenditure"],:] +
table5b[table5bRowDict["Total change in inventories"],:]);

#table 6
#importing relevant ASNA data
ASNAHouseInc = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204036_Household_Income_Account.xls", "Data1");
ASNANonFinInc = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204017_NonFin_Corp_Income_Account.xls", "Data1");
ASNAFinInc = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204025_Fin_Corp_Income_Account.xls", "Data1");
ASNAGovInc = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204030_GenGov_Income_Account.xls", "Data1");
ASNAExtInc = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204043_External_Accounts.xls", "Data1");
#initialising table
tableName = ["Households", "Non-Financial Corporations", "Financial Corporations", "General Government", "External", "Total"];
tableDict = Dict(tableName .=> [1:1:length(tableName);]);
table6 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table6[tableDict["Total"],tableDict["Households"]] = (first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Interest", x), string.(ASNAHouseInc[1,:]))])
+first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Imputed interest", x), string.(ASNAHouseInc[1,:]))]));
table6[tableDict["Total"],tableDict["Non-Financial Corporations"]] = (first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Interest", x), string.(ASNANonFinInc[1,:]))])
+first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("Property income attributed to insurance policyholders", x), string.(ASNANonFinInc[1,:]))]));
table6[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Interest", x), string.(ASNAFinInc[1,:]))]);
table6[tableDict["Total"],tableDict["General Government"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income receivable - Interest ;", x), string.(ASNAGovInc[1,:]))]);
table6[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Interest", x), string.(ASNAExtInc[1,:]))]);

table6[tableDict["Households"],tableDict["Total"]] = sum(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("Property income payable - Interest", x), string.(ASNAHouseInc[1,:]))]);
table6[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("Property income payable - Interest", x), string.(ASNANonFinInc[1,:]))]);
table6[tableDict["Financial Corporations"],tableDict["Total"]] = (first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("Property income payable - Interest", x), string.(ASNAFinInc[1,:]))])
+first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Property income attributed to insurance policy holders", x), string.(ASNAFinInc[1,:]))]));
table6[tableDict["General Government"],tableDict["Total"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income payable - Total interest ;", x), string.(ASNAGovInc[1,:]))]);
table6[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("Property income payable - Interest", x), string.(ASNAExtInc[1,:]))]);

if 0.98*sum(table6[:,length(tableName)])<sum(table6[length(tableName),:])<1.02*sum(table6[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 6")
end

#assuming gov only receives interest from fin corps
table6[tableDict["Financial Corporations"], tableDict["General Government"]] = table6[tableDict["Total"], tableDict["General Government"]];
#from other ASNA data
table6[tableDict["Financial Corporations"], tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("attributed", x), string.(ASNANonFinInc[1,:]))]);
table6[tableDict["Financial Corporations"], tableDict["Households"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("attributed", x), string.(ASNAFinInc[1,:]))])-first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("attributed", x), string.(ASNANonFinInc[1,:]))]);
table6[tableDict["General Government"], tableDict["Households"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income payable - Interest - On unfunded superannuation liabilities ;", x), string.(ASNAGovInc[1,:]))]);

#solve for missing values with scaling
table6Step3 = zeros(length(tableName),length(tableName));
table6Step3Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"],tableDict["General Government"]];
table6Step3Col = [tableDict["Households"],tableDict["External"]];
for i in table6Step3Col;
    for ring in table6Step3Row;
        table6Step3[ring,i] = (table6[tableDict["Total"],i]-sum(table6[1:(length(tableName)-1),i]))*(
            table6[ring,tableDict["Total"]]-sum(table6[ring,1:(length(tableName)-1)]))/sum(table6[
            table6Step3Row,tableDict["Total"]]-sum(eachcol(table6[table6Step3Row,1:(length(tableName)-1)])));
    end
end
table6 = table6+table6Step3;

table6Step4 = zeros(length(tableName),length(tableName));
table6Step4Row = [1:1:(length(tableName)-1);];
table6Step4Col = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"]];
for i in table6Step4Col;
    for ring in table6Step4Row;
        table6Step4[ring,i] = (table6[tableDict["Total"],i]-sum(table6[1:(length(tableName)-1),i]))*(
            table6[ring,tableDict["Total"]]-sum(table6[ring,1:(length(tableName)-1)]))/sum(table6[
            table6Step4Row,tableDict["Total"]]-sum(eachcol(table6[table6Step4Row,1:(length(tableName)-1)])));
    end
end
table6 = table6+table6Step4;

#table 7
#initialising table
table7 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table7[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Dividends", x), string.(ASNAHouseInc[1,:]))]);
table7[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Dividends", x), string.(ASNANonFinInc[1,:]))]);
table7[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Dividends", x), string.(ASNAFinInc[1,:]))]);
table7[tableDict["Total"],tableDict["General Government"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income receivable - Dividends ;", x), string.(ASNAGovInc[1,:]))]);
table7[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Dividends", x), string.(ASNAExtInc[1,:]))]);

table7[tableDict["Households"],tableDict["Total"]] = 0.0
table7[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Dividends", x), string.(ASNANonFinInc[1,:]))]);
table7[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Dividends", x), string.(ASNAFinInc[1,:]))]);
table7[tableDict["General Government"],tableDict["Total"]] = 0.0;
table7[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("payable - Dividends", x), string.(ASNAExtInc[1,:]))]);

if 0.98*sum(table7[:,length(tableName)])<sum(table7[length(tableName),:])<1.02*sum(table7[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 7")
end

#assuming gov only receives dividends from non-fin corps
table7[tableDict["Non-Financial Corporations"], tableDict["General Government"]] = table7[tableDict["Total"], tableDict["General Government"]];

#solve for missing values with scaling
table7Step3 = zeros(length(tableName),length(tableName));
table7Step3Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"]];
table7Step3Col = [tableDict["External"]];
for i in table7Step3Col;
    for ring in table7Step3Row;
        table7Step3[ring,i] = (table7[tableDict["Total"],i]-sum(table7[1:(length(tableName)-1),i]))*(
            table7[ring,tableDict["Total"]]-sum(table7[ring,1:(length(tableName)-1)]))/sum(table7[
            table7Step3Row,tableDict["Total"]]-sum(eachcol(table7[table7Step3Row,1:(length(tableName)-1)])));
    end
end
table7 = table7+table7Step3;

table7Step4 = zeros(length(tableName),length(tableName));
table7Step4Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"], tableDict["External"]];
table7Step4Col = [tableDict["Households"],tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"]];
for i in table7Step4Col;
    for ring in table7Step4Row;
        table7Step4[ring,i] = (table7[tableDict["Total"],i]-sum(table7[1:(length(tableName)-1),i]))*(
            table7[ring,tableDict["Total"]]-sum(table7[ring,1:(length(tableName)-1)]))/sum(table7[
            table7Step4Row,tableDict["Total"]]-sum(eachcol(table7[table7Step4Row,1:(length(tableName)-1)])));
    end
end
table7 = table7+table7Step4;

#table 8
#initialising table
table8 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table8[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNAHouseInc[1,:]))]);
table8[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNANonFinInc[1,:]))]);
table8[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNAFinInc[1,:]))]);
table8[tableDict["Total"],tableDict["General Government"]] = 0.0;
table8[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Reinvested", x), string.(ASNAExtInc[1,:]))]);

table8[tableDict["Households"],tableDict["Total"]] = 0.0
table8[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Reinvested", x), string.(ASNANonFinInc[1,:]))]);
table8[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Reinvested", x), string.(ASNAFinInc[1,:]))]);
table8[tableDict["General Government"],tableDict["Total"]] = 0.0;
table8[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("payable - Reinvested", x), string.(ASNAExtInc[1,:]))]);

if 0.98*sum(table8[:,length(tableName)])<sum(table8[length(tableName),:])<1.02*sum(table8[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 8")
end
#assuming that values can be negative
#solve missing values with Ipopt
println("ras for table 8")
mod8 = Model(Ipopt.Optimizer);
@variable(mod8, x[1:3, 1:4]>=0);
@NLobjective(mod8, Min, sum((x[i,j]) ^ 2 for i in 1:3, j in 1:4));
@constraint(mod8, sum(x[:,1]) == table8[tableDict["Total"],1]);
@constraint(mod8, sum(x[:,2]) == table8[tableDict["Total"],2]);
@constraint(mod8, sum(x[:,3]) == table8[tableDict["Total"],3]);
@constraint(mod8, sum(x[:,4]) == table8[tableDict["Total"],5]);
@constraint(mod8, sum(x[1,:]) == table8[2,tableDict["Total"]]);
@constraint(mod8, sum(x[2,:]) == table8[3,tableDict["Total"]]);
@constraint(mod8, sum(x[3,:]) == table8[5,tableDict["Total"]]);
#@constraint(mod8, x[3,4] == 0);
optimize!(mod8);
y= value.(x)
#table8b[[2,3,5], []] = table8[1:(length(tableName)-1), 1:(length(tableName)-1)] + value.(x);
#println(table8)

#=spread the external receivable totals between fin and non-fin Corporations
table8Step3 = zeros(length(tableName),length(tableName));
table8Step3Row = [tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"]];
table8Step3Col = [tableDict["External"]];
for i in table8Step3Col;
    for ring in table8Step3Row;
        table8Step3[ring,i] = (table8[tableDict["Total"],i]-sum(table8[1:(length(tableName)-1),i]))*(
            table8[ring,tableDict["Total"]]-sum(table8[ring,1:(length(tableName)-1)]))/sum(table8[
            table8Step3Row,tableDict["Total"]]-sum(eachcol(table8[table8Step3Row,1:(length(tableName)-1)])));
    end
end
table8 = table8+table8Step3;
=#

#table 9
#initialising table
table9 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table9[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Rent on natural", x), string.(ASNAHouseInc[1,:]))]);
table9[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Rent on natural", x), string.(ASNANonFinInc[1,:]))]);
table9[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Rent on natural", x), string.(ASNAFinInc[1,:]))]);
table9[tableDict["Total"],tableDict["General Government"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Property income receivable - Rent on natural assets ;", x), string.(ASNAGovInc[1,:]))]);
table9[tableDict["Total"],tableDict["External"]] = 0.0;

table9[tableDict["Households"],tableDict["Total"]] =  first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("payable - Rent on natural", x), string.(ASNAHouseInc[1,:]))]);
table9[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Rent on natural", x), string.(ASNANonFinInc[1,:]))]);
table9[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Rent on natural", x), string.(ASNAFinInc[1,:]))]);
table9[tableDict["General Government"],tableDict["Total"]] =  0.0;
table9[tableDict["External"],tableDict["Total"]] = 0.0;

if 0.98*sum(table9[:,length(tableName)])<sum(table9[length(tableName),:])<1.02*sum(table9[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 9")
end

#solve missing values with Ipopt
mod9 = Model(Ipopt.Optimizer);
@variable(mod9, x[1:(length(tableName)-1), 1:(length(tableName)-1)]>=0);
@NLobjective(mod9, Min, sum((x[i,j]) ^ 2 for i in 1:(length(tableName)-1), j in 1:(length(tableName)-1)));
for i in 1:(length(tableName)-1);
    @constraint(mod9, sum(x[:,i]) == table9[tableDict["Total"],i]-sum(table9[1:(length(tableName)-1),i]));
    @constraint(mod9, sum(x[i,:]) == table9[i,tableDict["Total"]]-sum(table9[i,1:(length(tableName)-1)]));
end;
@constraint(mod9, x[tableDict["Households"],tableDict["Households"]] == 0);
@constraint(mod9, x[tableDict["Households"],tableDict["Non-Financial Corporations"]] == 0);
optimize!(mod9);
table9[1:(length(tableName)-1), 1:(length(tableName)-1)] = table9[1:(length(tableName)-1), 1:(length(tableName)-1)] + value.(x);

#table 10
#initialising table
table10 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table10[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("Social assistance benefits", x), string.(ASNAHouseInc[1,:]))]);
table10[tableDict["General Government"],tableDict["Total"]] =  table10[tableDict["Total"],tableDict["Households"]];
#filling in missing values
table10[tableDict["General Government"],tableDict["Households"]] =  table10[tableDict["Total"],tableDict["Households"]];

#table 11
#initialising table
table11 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
#assuming that the large discrepancy in the paid and received claims is a result of households not reporting small claims
#table11[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Non-life", x), string.(ASNAHouseInc[1,:]))]);
table11[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Non-life", x), string.(ASNANonFinInc[1,:]))]);
table11[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Non-life", x), string.(ASNAFinInc[1,:]))]);
table11[tableDict["Total"],tableDict["Households"]]= table11[tableDict["Financial Corporations"],tableDict["Total"]] - table11[tableDict["Total"],tableDict["Non-Financial Corporations"]];
if 0.98*sum(table11[:,length(tableName)])<sum(table11[length(tableName),:])<1.02*sum(table11[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 11")
end
#filling in missing values
table11[tableDict["Financial Corporations"],tableDict["Households"]]=table11[tableDict["Total"],tableDict["Households"]];
table11[tableDict["Financial Corporations"],tableDict["Non-Financial Corporations"]] = table11[tableDict["Total"],tableDict["Non-Financial Corporations"]];

#table 12
#initialising table
table12 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
#also assuming in the same manner as in table 11 that the missing data is from the household side
table12[tableDict["Total"],tableDict["Households"]] = 0.0;
table12[tableDict["Total"],tableDict["Non-Financial Corporations"]] = 0.0;
table12[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Net non-life insurance premiums", x), string.(ASNAFinInc[1,:]))]);
table12[tableDict["Total"],tableDict["General Government"]] = 0.0;
table12[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Non-life insurance transfers", x), string.(ASNAExtInc[1,:]))]);

#table12[tableDict["Households"],tableDict["Total"]] =  first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("payable - Net non-life insurance premiums", x), string.(ASNAHouseInc[1,:]))]);
table12[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Net non-life insurance premiums", x), string.(ASNANonFinInc[1,:]))]);
table12[tableDict["Financial Corporations"],tableDict["Total"]] = 0.0;
table12[tableDict["General Government"],tableDict["Total"]] =  0.0;
table12[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("payable - Non-life insurance transfers", x), string.(ASNAExtInc[1,:]))]);
table12[tableDict["Households"],tableDict["Total"]] = sum(table12[length(tableName),:])-sum(table12[:,length(tableName)]);

if 0.98*sum(table12[:,length(tableName)])<sum(table12[length(tableName),:])<1.02*sum(table12[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 12")
end

#filling in empty values
table12[tableDict["External"],tableDict["Financial Corporations"]] = table12[tableDict["External"],tableDict["Total"]];
#=
mod12 = Model(Ipopt.Optimizer);
@variable(mod12, x[1:(length(tableName)-1), 1:(length(tableName)-1)]>=0);
@NLobjective(mod12, Min, sum((x[i,j])^2 for i in 1:(length(tableName)-1), j in 1:(length(tableName)-1)));
for i in 1:(length(tableName)-1);
    @constraint(mod12, sum(x[:,i]) == table12[tableDict["Total"],i]-sum(table12[1:(length(tableName)-1),i]));
    @constraint(mod12, sum(x[i,:]) == table12[i,tableDict["Total"]]-sum(table12[i,1:(length(tableName)-1)]));
end;
optimize!(mod12);
table12[1:(length(tableName)-1), 1:(length(tableName)-1)] = table12[1:(length(tableName)-1), 1:(length(tableName)-1)] + value.(x);
=#
#spread the external receivable totals between fin and non-fin Corporations
table12Step3 = zeros(length(tableName),length(tableName));
table12Step3Row = [tableDict["Non-Financial Corporations"],tableDict["Households"]];
table12Step3Col = [tableDict["External"],tableDict["Financial Corporations"]];
for i in table12Step3Col;
    for ring in table12Step3Row;
        table12Step3[ring,i] = (table12[tableDict["Total"],i]-sum(table12[1:(length(tableName)-1),i]))*(
            table12[ring,tableDict["Total"]]-sum(table12[ring,1:(length(tableName)-1)]))/sum(table12[
            table12Step3Row,tableDict["Total"]]-sum(eachcol(table12[table12Step3Row,1:(length(tableName)-1)])));
    end
end
table12 = table12+table12Step3;

#table 13
#initialising table
table13 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
table13[tableDict["Total"],tableDict["Households"]] = (first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Total other current transfers ;", x), string.(ASNAHouseInc[1,:]))])
+first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("receivable - Current transfers to ", x), string.(ASNAHouseInc[1,:]))]));
table13[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("receivable - Other current transfers ;", x), string.(ASNANonFinInc[1,:]))]);
table13[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("receivable - Other current transfers ;", x), string.(ASNAFinInc[1,:]))]);
table13[tableDict["Total"],tableDict["General Government"]] = first(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Secondary income receivable - Other current transfers ;", x), string.(ASNAGovInc[1,:]))]);
table13[tableDict["Total"],tableDict["External"]] = (first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Other current transfers ;", x), string.(ASNAExtInc[1,:]))])
+first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("receivable - Current international cooperation ;", x), string.(ASNAExtInc[1,:]))]));

table13[tableDict["Households"],tableDict["Total"]] =  first(ASNAHouseInc[ASNAYearRow,findall(x -> occursin("payable - Total other current transfers ;", x), string.(ASNAHouseInc[1,:]))]);
table13[tableDict["Non-Financial Corporations"],tableDict["Total"]] = (first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Current transfers to non-profit institutions ;", x), string.(ASNANonFinInc[1,:]))])
+ first(ASNANonFinInc[ASNAYearRow,findall(x -> occursin("payable - Other current transfers ;", x), string.(ASNANonFinInc[1,:]))]));
table13[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinInc[ASNAYearRow,findall(x -> occursin("payable - Other current transfers ;", x), string.(ASNAFinInc[1,:]))]);
table13[tableDict["General Government"],tableDict["Total"]] =  sum(ASNAGovInc[ASNAYearRow,findall(x -> occursin("General government ;  Secondary income payable - Other current transfers -", x), string.(ASNAGovInc[1,:]))]);
table13[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("payable - Other current transfers", x), string.(ASNAExtInc[1,:]))]);


if 0.98*sum(table13[:,length(tableName)])<sum(table13[length(tableName),:])<1.02*sum(table13[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 13")
end

#filling in empty values
table13Step3 = zeros(length(tableName),length(tableName));
table13Step3Row = [tableDict["Non-Financial Corporations"],[tableDict["Financial Corporations"],tableDict["General Government"],tableDict["External"]];
table13Step3Col = [tableDict["General Government"]];
for i in table13Step3Col;
    for ring in table13Step3Row;
        table13Step3[ring,i] = (table13[tableDict["Total"],i]-sum(table13[1:(length(tableName)-1),i]))*(
            table13[ring,tableDict["Total"]]-sum(table13[ring,1:(length(tableName)-1)]))/sum(table13[
            table13Step3Row,tableDict["Total"]]-sum(eachcol(table13[table13Step3Row,1:(length(tableName)-1)])));
    end
end
table13 = table13+table13Step3;

table13Step3 = zeros(length(tableName),length(tableName));
table13Step3Row = [tableDict["Households"],tableDict["Non-Financial Corporations"],tableDict["Financial Corporations"],tableDict["General Government"],tableDict["External"]];
table13Step3Col = [tableDict["Households"],tableDict["Non-Financial Corporations"],tableDict["External"]];
for i in table13Step3Col;
    for ring in table13Step3Row;
        table13Step3[ring,i] = (table13[tableDict["Total"],i]-sum(table13[1:(length(tableName)-1),i]))*(
            table13[ring,tableDict["Total"]]-sum(table13[ring,1:(length(tableName)-1)]))/sum(table13[
            table13Step3Row,tableDict["Total"]]-sum(eachcol(table13[table13Step3Row,1:(length(tableName)-1)])));
    end
end
table13 = table13+table13Step3;

#table 14
#initialising table
table14 = zeros(length(tableName),length(tableName));
#filling in data from ASNA
table14[tableDict["General Government"],tableDict["Households"]] = first(ASNAHouseCap[ASNAYearRow,findall(x -> occursin("Capital transfers, receivable from general government ;", x), string.(ASNAHouseCap[1,:]))]);
table14[tableDict["Total"],tableDict["Households"]] = first(ASNAHouseCap[ASNAYearRow,findall(x -> occursin("Other capital transfers, receivable ;", x), string.(ASNAHouseCap[1,:]))])+table14[tableDict["General Government"],tableDict["Households"]];
table14[tableDict["General Government"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinCap[ASNAYearRow,findall(x -> occursin("Capital transfers, receivable from general government ;", x), string.(ASNANonFinCap[1,:]))]);
table14[tableDict["Total"],tableDict["Non-Financial Corporations"]] = first(ASNANonFinCap[ASNAYearRow,findall(x -> occursin("Other capital transfers, receivable ;", x), string.(ASNANonFinCap[1,:]))])+table14[tableDict["General Government"],tableDict["Non-Financial Corporations"]];
table14[tableDict["General Government"],tableDict["Financial Corporations"]] = first(ASNAFinCap[ASNAYearRow,findall(x -> occursin("Capital transfers, receivable from general government ;", x), string.(ASNAFinCap[1,:]))]);
table14[tableDict["Total"],tableDict["Financial Corporations"]] = first(ASNAFinCap[ASNAYearRow,findall(x -> occursin("Other capital transfers, receivable ;", x), string.(ASNAFinCap[1,:]))])+table14[tableDict["General Government"],tableDict["Financial Corporations"]];
table14[tableDict["Households"],tableDict["General Government"]] = first(ASNAHouseCap[ASNAYearRow,findall(x -> occursin("Capital transfers, payable to general government ;", x), string.(ASNAHouseCap[1,:]))]);
table14[tableDict["Non-Financial Corporations"],tableDict["General Government"]] = first(ASNANonFinCap[ASNAYearRow,findall(x -> occursin("Capital transfers, payable to general government ;", x), string.(ASNANonFinCap[1,:]))]);
table14[tableDict["Financial Corporations"],tableDict["General Government"]] = first(ASNAFinCap[ASNAYearRow,findall(x -> occursin("Capital transfers, payable to general government ;", x), string.(ASNAFinCap[1,:]))]);
table14[tableDict["Total"],tableDict["General Government"]] = first(ASNAGovCap[ASNAYearRow,findall(x -> occursin("General government ;  Capital transfers, receivable ;", x), string.(ASNAGovCap[1,:]))]);
table14[tableDict["Total"],tableDict["External"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("Capital transfers, receivable ;", x), string.(ASNAExtInc[1,:]))]);
table14[tableDict["Households"],tableDict["Total"]] = first(ASNAHouseCap[ASNAYearRow,findall(x -> occursin("Other capital transfers, payable ;", x), string.(ASNAHouseCap[1,:]))])+table14[tableDict["Households"],tableDict["General Government"]];
table14[tableDict["Non-Financial Corporations"],tableDict["Total"]] = first(ASNANonFinCap[ASNAYearRow,findall(x -> occursin("Other capital transfers, payable ;", x), string.(ASNANonFinCap[1,:]))])+table14[tableDict["Non-Financial Corporations"],tableDict["General Government"]];
table14[tableDict["Financial Corporations"],tableDict["Total"]] = first(ASNAFinCap[ASNAYearRow,findall(x -> occursin("Other capital transfers, payable ;", x), string.(ASNAFinCap[1,:]))])+table14[tableDict["Financial Corporations"],tableDict["General Government"]];
table14[tableDict["General Government"],tableDict["Total"]] = first(ASNAGovCap[ASNAYearRow,findall(x -> occursin("General government ;  Capital transfers, payable ;", x), string.(ASNAGovCap[1,:]))]);
table14[tableDict["External"],tableDict["Total"]] = first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("Capital transfers, payable ;", x), string.(ASNAExtInc[1,:]))]);
if 0.98*sum(table14[:,length(tableName)])<sum(table14[length(tableName),:])<1.02*sum(table14[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 14")
end

#solve missing values with Ipopt
mod14 = Model(Ipopt.Optimizer);
@variable(mod14, x[1:(length(tableName)-1), 1:(length(tableName)-1)]>=0);
@NLobjective(mod14, Min, sum((x[i,j]) ^ 2 for i in 1:(length(tableName)-1), j in 1:(length(tableName)-1)));
for i in 1:(length(tableName)-1);
    @constraint(mod14, sum(x[:,i]) == table14[tableDict["Total"],i]-sum(table14[1:(length(tableName)-1),i]));
    @constraint(mod14, sum(x[i,:]) == table14[i,tableDict["Total"]]-sum(table14[i,1:(length(tableName)-1)]));
    @constraint(mod14, x[i,i] == 0);
end;

optimize!(mod14);
table14[1:(length(tableName)-1), 1:(length(tableName)-1)] = table14[1:(length(tableName)-1), 1:(length(tableName)-1)] + value.(x);

#table 15
#initialising table
table15 = zeros(length(tableName),length(tableName));
#allocating total collumn and row data
if first(ASNAHouseCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAHouseCap[1,:]))])<=0
    table15[tableDict["Households"],tableDict["Total"]] = abs(first(ASNAHouseCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAHouseCap[1,:]))]));
else
    table15[tableDict["Total"],tableDict["Households"]] = abs(first(ASNAHouseCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAHouseCap[1,:]))]));
end
if first(ASNANonFinCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNANonFinCap[1,:]))])<=0
    table15[tableDict["Non-Financial Corporations"],tableDict["Total"]] = abs(first(ASNANonFinCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNANonFinCap[1,:]))]));
else
    table15[tableDict["Total"],tableDict["Non-Financial Corporations"]] = abs(first(ASNANonFinCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNANonFinCap[1,:]))]));
end
if first(ASNAFinCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAFinCap[1,:]))])<=0
    table15[tableDict["Financial Corporations"],tableDict["Total"]] = abs(first(ASNAFinCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAFinCap[1,:]))]));
else
    table15[tableDict["Total"],tableDict["Financial Corporations"]] = abs(first(ASNAFinCap[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAFinCap[1,:]))]));
end
if first(ASNAGovCap[ASNAYearRow,findall(x -> occursin("General government ;  Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAGovCap[1,:]))])<=0
    table15[tableDict["General Government"],tableDict["Total"]] = abs(first(ASNAGovCap[ASNAYearRow,findall(x -> occursin("General government ;  Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAGovCap[1,:]))]));
else
    table15[tableDict["Total"],tableDict["General Government"]] = abs(first(ASNAGovCap[ASNAYearRow,findall(x -> occursin("General government ;  Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAGovCap[1,:]))]));
end
if first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAExtInc[1,:]))])<=0
    table15[tableDict["External"],tableDict["Total"]] = abs(first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAExtInc[1,:]))]));
else
    table15[tableDict["Total"],tableDict["External"]] = abs(first(ASNAExtInc[ASNAYearRow,findall(x -> occursin("Acquisitions less disposals of non-produced non-financial assets ;", x), string.(ASNAExtInc[1,:]))]));
end

if 0.98*sum(table15[:,length(tableName)])<sum(table15[length(tableName),:])<1.02*sum(table15[:,length(tableName)])
else
    error("Large discrepancy in row and collumn total sums in table 15")
end

#solve missing values with Ipopt
mod15 = Model(Ipopt.Optimizer);
@variable(mod15, x[1:(length(tableName)-1), 1:(length(tableName)-1)]>=0);
@NLobjective(mod15, Min, sum((x[i,j]) ^ 2 for i in 1:(length(tableName)-1), j in 1:(length(tableName)-1)));
for i in 1:(length(tableName)-1);
    @constraint(mod15, sum(x[:,i]) == table15[tableDict["Total"],i]-sum(table15[1:(length(tableName)-1),i]));
    @constraint(mod15, sum(x[i,:]) == table15[i,tableDict["Total"]]-sum(table15[i,1:(length(tableName)-1)]));
    @constraint(mod15, x[i,i] == 0);
end;

optimize!(mod15);
table15[1:(length(tableName)-1), 1:(length(tableName)-1)] = table15[1:(length(tableName)-1), 1:(length(tableName)-1)] + value.(x);

#table 16a
#initialising table
table16aNameRow = ["Netsaving;", "Totalnetcapitaltransfers;", "Grossfixedcapitalformation;", "Netlending(+)/netborrowing(-);"];
table16aDataCol = [ASNAHouseCap, ASNANonFinCap, ASNAFinCap, ASNAGovCap, ASNAExtInc];
table16aRowDict = Dict(table16aNameRow .=> [1:1:length(table16aNameRow);]);
table16aDataDict = Dict(tableName[1:length(tableName)-1] .=> table16aDataCol);
table16a = zeros(length(table16aNameRow),length(tableName));
#filling in values
for i in tableName
    if i == "Total"
        for ring in table16aNameRow
            table16a[table16aRowDict[ring],tableDict[i]]=sum(table16a[table16aRowDict[ring],:]);
        end
    elseif i == "General Government"
        for ring in table16aNameRow
            table16a[table16aRowDict[ring],tableDict[i]]=first(table16aDataDict[i][ASNAYearRow,findall(x -> occursin("Generalgovernment;"*ring, x), filter.(x -> !isspace(x), string.(table16aDataDict[i][1,:])))]);
        end
    else
        for ring in table16aNameRow
            if isempty(findall(x -> occursin(ring, x), filter.(x -> !isspace(x),string.(table16aDataDict[i][1,:]))))
                table16a[table16aRowDict[ring],tableDict[i]]=0.0;
            else
                table16a[table16aRowDict[ring],tableDict[i]]=first(table16aDataDict[i][ASNAYearRow,findall(x -> occursin(ring, x), filter.(x -> !isspace(x), string.(table16aDataDict[i][1,:])))]);
            end
        end
    end
end

#table 16b
#importing relevant data from ASNA
#importing relevant ASNA data
ASNAHouseFin = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204040_Household_Financial_Account.xls", "Data1");
ASNANonFinFin = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204019_NonFin_Corp_Financial_Account.xls", "Data1");
ASNAFinFin = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204027_Fin_Corp_Financial_Account.xls", "Data1");
ASNAGovFin = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204033_GenGov_Financial_Account.xls", "Data1");
ASNAExtFin = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204044_External_Financial_Account.xls", "Data1");
#initialising table
#make no spaces
table16bNameRow = ["Acquisitionoffinancialassets-MonetarygoldandSDRs;","Acquisitionoffinancialassets-Currencyanddeposits;",
"Acquisitionoffinancialassets-Billsofexchange;","Acquisitionoffinancialassets-Onenamepaper–issuedinAustralia;",
"Acquisitionoffinancialassets-Onenamepaper–issuedoffshore;","Acquisitionoffinancialassets-Bondsetc.–issuedinAustralia;",
"Acquisitionoffinancialassets-Bondsetc.–issuedoffshore;","Acquisitionoffinancialassets-Derivatives;",
"Acquisitionoffinancialassets-Loansandplacements;","Acquisitionoffinancialassets-Sharesandotherequity;",
"Acquisitionoffinancialassets-Totalinsurancetechnicalreserves;","Acquisitionoffinancialassets-Otheraccounts",
"Incurrenceofliabilities-MonetarygoldandSDRs;","Incurrenceofliabilities-Currencyanddeposits;",
"Incurrenceofliabilities-Billsofexchange;","Incurrenceofliabilities-Onenamepaper–issuedinAustralia;",
"Incurrenceofliabilities-Onenamepaper–issuedoffshore;","Incurrenceofliabilities-Bondsetc.–issuedinAustralia;",
"Incurrenceofliabilities-Bondsetc.–issuedoffshore;","Incurrenceofliabilities-Derivatives;",
"Incurrenceofliabilities-Loansandplacements;","Incurrenceofliabilities-Sharesandotherequity;",
"Incurrenceofliabilities-Totalinsurancetechnicalreserves;","Incurrenceofliabilities-Otheraccounts"];
table16bDataCol = [ASNAHouseFin, ASNANonFinFin, ASNAFinFin, ASNAGovFin, ASNAExtFin];
table16bRowDict = Dict(table16bNameRow .=> [1:1:length(table16bNameRow);]);
table16bDataDict = Dict(tableName[1:length(tableName)-1] .=> table16bDataCol);
table16b = zeros(length(table16bNameRow),length(tableName));
#filling in values
for i in tableName
    if i == "Total"
        for ring in table16bNameRow
            table16b[table16bRowDict[ring],tableDict[i]]=sum(table16b[table16bRowDict[ring],:]);
        end
    elseif i == "General Government"
        for ring in table16bNameRow
            if isempty(findall(x -> occursin("Generalgovernment-"*ring, x), filter.(x -> !isspace(x), string.(table16bDataDict[i][1,:]))))
                table16b[table16bRowDict[ring],tableDict[i]]=0.0;
            else
                table16b[table16bRowDict[ring],tableDict[i]]=1000*first(table16bDataDict[i][findall(x -> occursin("2019", x), string.(table16bDataDict[i][:,1])),findall(x -> occursin("Generalgovernment-"*ring, x), filter.(x -> !isspace(x), string.(table16bDataDict[i][1,:])))]);
            end
        end
    elseif i == "External"
        for ring in table16bNameRow
            key = ring
            if occursin("liabilities",ring) && occursin("issuedoffshore",ring)
                ring = replace(ring,"–issuedoffshore" => "")
            else
            end
            if isempty(findall(x -> occursin(ring, x), filter.(x -> !isspace(x), string.(table16bDataDict[i][1,:]))))
                table16b[table16bRowDict[key],tableDict[i]]=0.0;
            else
                table16b[table16bRowDict[key],tableDict[i]]=1000*first(table16bDataDict[i][findall(x -> occursin("2019", x), string.(table16bDataDict[i][:,1])),findall(x -> occursin(ring, x), filter.(x -> !isspace(x), string.(table16bDataDict[i][1,:])))]);
            end
        end
    else
        for ring in table16bNameRow
            if isempty(findall(x -> occursin(ring, x), filter.(x -> !isspace(x), string.(table16bDataDict[i][1,:]))))
                table16b[table16bRowDict[ring],tableDict[i]]=0.0;
            else
                table16b[table16bRowDict[ring],tableDict[i]]=1000*first(table16bDataDict[i][findall(x -> occursin("2019", x), string.(table16bDataDict[i][:,1])),findall(x -> occursin(ring, x), filter.(x -> !isspace(x), string.(table16bDataDict[i][1,:])))]);
            end
        end
    end
end

table16bAqcRow = findall(x -> occursin("Acquisition", x), string.(table16bNameRow));
table16bLiaRow = findall(x -> occursin("liabilities", x), string.(table16bNameRow));
table16bAqcTotal = sum(eachrow(table16b[table16bAqcRow,:]));
table16bLiaTotal = sum(eachrow(table16b[table16bLiaRow,:]));
table16SurplusOrDeficit = table16bAqcTotal - table16bLiaTotal;

#initialising table
#make no spaces
eAndONameRow = ["Neterrorsandomissions;"];
eAndODataCol = [ASNAHouseFin, ASNANonFinFin, ASNAFinFin, ASNAGovFin, ASNAExtFin];
eAndORowDict = Dict(eAndONameRow .=> [1:1:length(eAndONameRow);]);
eAndODataDict = Dict(tableName[1:length(tableName)-1] .=> eAndODataCol);
eAndO = zeros(length(eAndONameRow),length(tableName));
#filling in values
for i in tableName
    if i == "Total"
        for ring in eAndONameRow
            eAndO[eAndORowDict[ring],tableDict[i]]=sum(eAndO[eAndORowDict[ring],:]);
        end
    elseif i == "General Government"
        for ring in eAndONameRow
            if isempty(findall(x -> occursin("Generalgovernment-"*ring, x), filter.(x -> !isspace(x), string.(eAndODataDict[i][1,:]))))
                eAndO[eAndORowDict[ring],tableDict[i]]=0.0;
            else
                eAndO[eAndORowDict[ring],tableDict[i]]=1000*first(eAndODataDict[i][findall(x -> occursin("2019", x), string.(eAndODataDict[i][:,1])),findall(x -> occursin("Generalgovernment-"*ring, x), filter.(x -> !isspace(x), string.(eAndODataDict[i][1,:])))]);
            end
        end
    else
        for ring in eAndONameRow
            if isempty(findall(x -> occursin(ring, x), filter.(x -> !isspace(x), string.(eAndODataDict[i][1,:]))))
                eAndO[eAndORowDict[ring],tableDict[i]]=0.0;
            else
                eAndO[eAndORowDict[ring],tableDict[i]]=1000*first(eAndODataDict[i][findall(x -> occursin("2019", x), string.(eAndODataDict[i][:,1])),findall(x -> occursin(ring, x), filter.(x -> !isspace(x), string.(eAndODataDict[i][1,:])))]);
            end
        end
    end
end

table17TableNames = ["MonetarygoldandSDRs;","Currencyanddeposits;","Billsofexchange;",
"Onenamepaper–issuedinAustralia;","Onenamepaper–issuedoffshore;","Bondsetc.–issuedinAustralia;",
"Bondsetc.–issuedoffshore;","Derivatives;","Loansandplacements;","Sharesandotherequity;",
"Totalinsurancetechnicalreserves;","Otheraccounts"];

function myfind(c)
    a = similar(c, Int)
    count = 1
    @inbounds for i in eachindex(c)
        a[count] = i
        count += (c[i] != zero(eltype(c)))
    end
    return resize!(a, count-1)
end

table17=Vector{Union{Matrix{Float64}, Nothing}}(undef, length(table17TableNames));
for ring in [1:1:length(table17TableNames);]
    println("this round is number ", ring)
    table17[ring] = zeros(length(tableName),length(tableName));
    table17[ring][length(tableName),:] = table16b[table16bRowDict["Acquisitionoffinancialassets-"*table17TableNames[ring]],:];
    table17[ring][:,length(tableName)] = table16b[table16bRowDict["Incurrenceofliabilities-"*table17TableNames[ring]],:];
    if sum(table17[ring][length(tableName),1:(length(tableName)-1)]) == sum(table17[ring][1:(length(tableName)-1),length(tableName)])
    else
        table17RowTot = sum(table17[ring][length(tableName),1:(length(tableName)-1)]);
        table17ColTot = sum(table17[ring][1:(length(tableName)-1),length(tableName)]);
        table17TotAv = (table17RowTot + table17ColTot)/2;
        if 0.98*abs(table17RowTot)<abs(table17TotAv)<1.04*abs(table17RowTot)
        else
            error("Large discrepancy in row and collumn total sums in table 17."*string(ring))
        end
        table17[ring][length(tableName),length(tableName)] = table17TotAv;
        for i in 1:(length(tableName)-1);
            table17[ring][length(tableName),i] = table17[ring][length(tableName),i]*table17TotAv/table17RowTot;
            table17[ring][i,length(tableName)] = table17[ring][i,length(tableName)]*table17TotAv/table17ColTot;
        end
    end
    nonZerosRow = myfind(table17[ring][:,length(tableName)]);
    nonZerosCol = myfind(table17[ring][length(tableName),:]);
    if length(nonZerosRow)-1 == 1;
        table17[ring][nonZerosRow[1],nonZerosCol] = table17[ring][tableDict["Total"],nonZerosCol];
    elseif length(nonZerosCol)-1 == 1;
        table17[ring][nonZerosRow,nonZerosCol[1]] = table17[ring][nonZerosRow,tableDict["Total"]];
    else
        mod17 = Model(Ipopt.Optimizer);
        @variable(mod17, x[1:(length(nonZerosRow)-1), 1:(length(nonZerosCol)-1)]);
        @NLobjective(mod17, Min, sum((x[i,j]) ^ 2 for i in 1:(length(nonZerosRow)-1), j in 1:(length(nonZerosCol)-1)));
        for i in 1:length(nonZerosRow)-1;
            @constraint(mod17, sum(x[i,:]) == table17[ring][nonZerosRow[i],tableDict["Total"]]);
        end
        for i in 1:(length(nonZerosCol)-1);
            @constraint(mod17, sum(x[:,i]) == table17[ring][tableDict["Total"],nonZerosCol[i]]);
        end
        if tableDict["External"] in nonZerosRow && tableDict["External"] in nonZerosCol
            @constraint(mod17, x[length(nonZerosRow)-1,length(nonZerosCol)-1] == 0);
        else
        end
        optimize!(mod17);
        table17[ring][nonZerosRow[1:length(nonZerosRow)-1], nonZerosCol[1:length(nonZerosCol)-1]] = value.(x);
    end
end

table19 = sum(table17[:]);
table19TotRow = table19[:,tableDict["Total"]]+table16a[table16aRowDict["Netlending(+)/netborrowing(-);"],:];
table19EAndO = table19[tableDict["Total"],:]-table19TotRow;

#final bits and pieces to fill in gaps
ASNAIncomeFromGDP = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204006_Income_from_GDP.xls", "Data1");
productionToHouseholds = (first(ASNAIncomeFromGDP[findall(x -> occursin("2019", x), string.(ASNAIncomeFromGDP[:,1])),findall(x -> occursin("Compensationofemployees;", x), filter.(x -> !isspace(x), string.(ASNAIncomeFromGDP[1,:])))])
-first(ASNAExtInc[findall(x -> occursin("2019", x), string.(ASNAExtInc[:,1])),findall(x -> occursin("receivable-Compensationofemployees;", x), filter.(x -> !isspace(x), string.(ASNAExtInc[1,:])))])
+first(ASNAIncomeFromGDP[findall(x -> occursin("2019", x), string.(ASNAIncomeFromGDP[:,1])),findall(x -> occursin("Dwellingsownedbypersons;Grossoperatingsurplus;", x), filter.(x -> !isspace(x), string.(ASNAIncomeFromGDP[1,:])))])
+first(ASNAIncomeFromGDP[findall(x -> occursin("2019", x), string.(ASNAIncomeFromGDP[:,1])),findall(x -> occursin("Grossmixedincome;", x), filter.(x -> !isspace(x), string.(ASNAIncomeFromGDP[1,:])))]));

#current to current transfers
#from households
currentToCurrent = transpose(table6 + table7 + table8 + table9 + table10 + table11 + table12 + table13);
currentToCurrent[tableDict["General Government"],tableDict["Households"]] = (currentToCurrent[tableDict["General Government"],tableDict["Households"]] + first(ASNAHouseInc[findall(x -> occursin("2019", x), string.(ASNAHouseInc[:,1])),findall(x -> occursin("payable-Incometax;", x), filter.(x -> !isspace(x), string.(ASNAHouseInc[1,:])))])
+ first(ASNAHouseInc[findall(x -> occursin("2019", x), string.(ASNAHouseInc[:,1])),findall(x -> occursin("payable-Othercurrenttaxes", x), filter.(x -> !isspace(x), string.(ASNAHouseInc[1,:])))]) + IO[IORowDict["P3"], IOColDict["Q1"]]);
currentToCurrent[tableDict["External"],tableDict["Households"]] = currentToCurrent[tableDict["External"],tableDict["Households"]] + sum(IO[IORowDict["P5"]:IORowDict["P6"],IOColDict["Q1"]]);
#from non-financial
currentToCurrent[tableDict["General Government"],tableDict["Non-Financial Corporations"]] = currentToCurrent[tableDict["General Government"],tableDict["Non-Financial Corporations"]] + first(ASNANonFinInc[findall(x -> occursin("2019", x), string.(ASNANonFinInc[:,1])),findall(x -> occursin("Secondaryincomepayable-Currenttaxesonincome,wealth,etc-Total;", x), filter.(x -> !isspace(x), string.(ASNANonFinInc[1,:])))]);
#from financial
currentToCurrent[tableDict["General Government"],tableDict["Financial Corporations"]] = currentToCurrent[tableDict["General Government"],tableDict["Financial Corporations"]] + first(ASNAFinInc[findall(x -> occursin("2019", x), string.(ASNAFinInc[:,1])),findall(x -> occursin("Secondaryincomepayable-Currenttaxesonincome,wealth,etc-Total;", x), filter.(x -> !isspace(x), string.(ASNAFinInc[1,:])))]);
#from general Government
currentToCurrent[tableDict["External"],tableDict["General Government"]] = currentToCurrent[tableDict["External"],tableDict["General Government"]] + sum(IO[IORowDict["P5"]:IORowDict["P6"],IOColDict["Q2"]]);
#from external 
currentToCurrent[tableDict["General Government"],tableDict["External"]] = currentToCurrent[tableDict["General Government"],tableDict["External"]] + sum(IO[IORowDict["P3"],IOColDict["Q2"]]) + first(ASNAExtInc[findall(x -> occursin("2019", x), string.(ASNAExtInc[:,1])),findall(x -> occursin("payable-Currenttaxesonincome,wealth,etc;", x), filter.(x -> !isspace(x), string.(ASNAExtInc[1,:])))]);
currentToCurrent[tableDict["Households"],tableDict["External"]] = currentToCurrent[tableDict["Households"],tableDict["External"]] + first(ASNAExtInc[findall(x -> occursin("2019", x), string.(ASNAExtInc[:,1])),findall(x -> occursin("payable-Compensationofemployees;", x), filter.(x -> !isspace(x), string.(ASNAExtInc[1,:])))]);

#capital to capital transfers
capitalToCapital = table14 + table15 + table19

ASNAConsFixCap = ExcelReaders.readxlsheet("ASNAData"*pathmark*"5204047_Cons_Fixed_Capital_By_Industry.xls", "Data1");

samName = ["Production Activities", "Factors of Production", "Households Current Account", "Non-Financial Corporations Current Account",
"Financial Corporations Current Account", "Government Current Account", "Foreigners Current Account",
"Households Capital Account", "Non-Financial Corporations Capital Account", "Financial Corporations Capital Account",
"Government Capital Account", "Foreigners Capital Account", "Net Errors and Ommisions", "Total"];
samDict = Dict(samName .=> [1:1:length(samName);]);
sam = zeros(length(samName),length(samName));
sam[samDict["Production Activities"],samDict["Production Activities"]] = IO[IORowDict["T1"],IOColDict["T4"]];
sam[samDict["Production Activities"],samDict["Households Current Account"]] = IO[IORowDict["T1"],IOColDict["Q1"]];
sam[samDict["Production Activities"],samDict["Government Current Account"]] = IO[IORowDict["T1"],IOColDict["Q2"]];
sam[samDict["Production Activities"],samDict["Foreigners Current Account"]] = IO[IORowDict["T1"],IOColDict["Q7"]];
sam[samDict["Production Activities"],samDict["Households Capital Account"]:samDict["Government Capital Account"]] = table5c[table5cRowDict["Domestic Commodities"],table5cColDict["Households"]:table5cColDict["General Government"]];
sam[samDict["Factors of Production"], samDict["Production Activities"]] = sum(IO[IORowDict["P1"]:IORowDict["P2"],IOColDict["T4"]]);
sam[samDict["Households Current Account"], samDict["Factors of Production"]] = productionToHouseholds;
sam[samDict["Non-Financial Corporations Current Account"], samDict["Factors of Production"]] = first(ASNANonFinInc[findall(x -> occursin("2019", x), string.(ASNANonFinInc[:,1])),findall(x -> occursin("Grossoperatingsurplus;", x), filter.(x -> !isspace(x), string.(ASNANonFinInc[1,:])))]);
sam[samDict["Financial Corporations Current Account"], samDict["Factors of Production"]] = first(ASNAFinInc[findall(x -> occursin("2019", x), string.(ASNAFinInc[:,1])),findall(x -> occursin("Grossoperatingsurplus;", x), filter.(x -> !isspace(x), string.(ASNAFinInc[1,:])))]);
sam[samDict["Government Current Account"], samDict["Factors of Production"]] = first(ASNAGovInc[findall(x -> occursin("2019", x), string.(ASNAGovInc[:,1])),findall(x -> occursin("Generalgovernment;Consumptionoffixedcapital;", x), filter.(x -> !isspace(x), string.(ASNAGovInc[1,:])))]);
sam[samDict["Foreigners Current Account"], samDict["Factors of Production"]] = first(ASNAExtInc[findall(x -> occursin("2019", x), string.(ASNAExtInc[:,1])),findall(x -> occursin("receivable-Compensationofemployees;", x), filter.(x -> !isspace(x), string.(ASNAExtInc[1,:])))]);
sam[samDict["Government Current Account"], samDict["Production Activities"]] = sum(IO[IORowDict["P3"]:IORowDict["P4"],IOColDict["T4"]]);
sam[samDict["Foreigners Current Account"], samDict["Production Activities"]] = sum(IO[IORowDict["P5"]:IORowDict["P6"],IOColDict["T4"]]);
sam[samDict["Households Current Account"]:samDict["Foreigners Current Account"],samDict["Households Current Account"]:samDict["Foreigners Current Account"]]=currentToCurrent[1:length(tableName)-1,1:length(tableName)-1];
sam[samDict["Households Capital Account"], samDict["Households Current Account"]] = (first(ASNAHouseInc[findall(x -> occursin("2019", x), string.(ASNAHouseInc[:,1])),findall(x -> occursin("Netsaving;", x), filter.(x -> !isspace(x), string.(ASNAHouseInc[1,:])))])
+first(ASNAConsFixCap[findall(x -> occursin("2019", x), string.(ASNAConsFixCap[:,1])),findall(x -> occursin("Households;Dwellingsownedbypersons;", x), filter.(x -> !isspace(x), string.(ASNAConsFixCap[1,:])))])
+first(ASNAConsFixCap[findall(x -> occursin("2019", x), string.(ASNAConsFixCap[:,1])),findall(x -> occursin("Households;Other;", x), filter.(x -> !isspace(x), string.(ASNAConsFixCap[1,:])))]));
sam[samDict["Non-Financial Corporations Capital Account"], samDict["Non-Financial Corporations Current Account"]] = first(ASNANonFinInc[findall(x -> occursin("2019", x), string.(ASNANonFinInc[:,1])),findall(x -> occursin("Consumptionoffixedcapital;", x), filter.(x -> !isspace(x), string.(ASNANonFinInc[1,:])))])+first(ASNANonFinInc[findall(x -> occursin("2019", x), string.(ASNANonFinInc[:,1])),findall(x -> occursin("Netsaving;", x), filter.(x -> !isspace(x), string.(ASNANonFinInc[1,:])))]);
sam[samDict["Financial Corporations Capital Account"], samDict["Financial Corporations Current Account"]] = first(ASNAFinInc[findall(x -> occursin("2019", x), string.(ASNAFinInc[:,1])),findall(x -> occursin("Consumptionoffixedcapital;", x), filter.(x -> !isspace(x), string.(ASNAFinInc[1,:])))])+first(ASNAFinInc[findall(x -> occursin("2019", x), string.(ASNAFinInc[:,1])),findall(x -> occursin("Netsaving;", x), filter.(x -> !isspace(x), string.(ASNAFinInc[1,:])))]);
sam[samDict["Government Capital Account"], samDict["Government Current Account"]] = first(ASNAGovInc[findall(x -> occursin("2019", x), string.(ASNAGovInc[:,1])),findall(x -> occursin("Generalgovernment;Consumptionoffixedcapital;", x), filter.(x -> !isspace(x), string.(ASNAGovInc[1,:])))])+first(ASNAGovInc[findall(x -> occursin("2019", x), string.(ASNAGovInc[:,1])),findall(x -> occursin("Generalgovernment;Netsaving;", x), filter.(x -> !isspace(x), string.(ASNAGovInc[1,:])))]);
sam[samDict["Foreigners Capital Account"], samDict["Foreigners Current Account"]] = first(ASNAExtInc[findall(x -> occursin("2019", x), string.(ASNAExtInc[:,1])),findall(x -> occursin("Balanceonexternalincomeaccount;", x), filter.(x -> !isspace(x), string.(ASNAExtInc[1,:])))]);

sam[samDict["Households Capital Account"]:samDict["Foreigners Capital Account"],samDict["Households Capital Account"]:samDict["Foreigners Capital Account"]]= capitalToCapital[1:length(tableDict)-1,1:length(tableDict)-1];
sam[samDict["Government Current Account"],samDict["Households Capital Account"]:samDict["Government Capital Account"]] = table5c[table5cRowDict["Taxes less subsidies on products"],1:length(table5cNameCol)-1] + table5c[table5cRowDict["Other taxes less subsidies on investment"],1:length(table5cNameCol)-1];
sam[samDict["Foreigners Current Account"],samDict["Households Capital Account"]:samDict["Government Capital Account"]] = table5c[table5cRowDict["Imported Commodities"],1:length(table5cNameCol)-1];
sam[samDict["Net Errors and Ommisions"],samDict["Production Activities"]] =  table19EAndO[tableDict["Total"]];
sam[samDict["Net Errors and Ommisions"],samDict["Households Capital Account"]:samDict["Foreigners Capital Account"]] = -1 .* table19EAndO[1:length(tableName)-1];

for i in 1:length(samName)-1
    sam[i,samDict["Total"]]=sum(sam[i,:]);
    sam[samDict["Total"],i]=sum(sam[:,i]);
end
#=convert dataframe to dictionary
function increment!( d::Dict{S, T}, k::S, i::T) where {T<:Real, S<:Any}
    if haskey(d, k)
        d[k] += i
    else
        d[k] = i
    end
end
increment!(d::Dict{S, T}, k::S ) where {T<:Real, S<:Any} = increment!( d, k, one(T))

function df2dict( df::DataFrame, key_col::Symbol, val_col::Symbol=:null)
    keytype = typeof(df[1,key_col])
    if val_col == :null
        valtype = Int
    else
        valtype = typeof(df[1,val_col])
    end
    D = Dict{keytype, valtype}()
    for i=1:size(df,1)
        if !ismissing(df[i,key_col])
            if val_col == :null
                increment!( D, df[i,key_col] )
            elseif valtype <: Real
                increment!( D, df[i,key_col], df[i,val_col] )
            else
                if haskey(D, df[i,key_col])
                    @warn("non-unique entry: $(df[i,key_col])")
                else
                    D[df[i,key_col]] = df[i,val_col]
                end
            end
        end
    end
    return D
end
df[!, "IOCode"]=IOcode
insertcols!(df, 2, :name => vector)
D = df2dict(df, :IOCode, :x3)
=#

#notes
#double check for missing accounts i.e. they said 0 for 96-97 so they were made to be 0 for this year, but may not be in other years
#also check for other missing accounts, confusingly labelled so skipped over etc
#implement loop and more ras-ing
#take out spaces in all find functions
#generalise the ASNAYearRow to just search every time
#rename tables from numbers to abbreviations