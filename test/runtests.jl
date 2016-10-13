using ExcelReaders
using BaseTestNext
using PyTest
using PyCall
using DataArrays
using DataFrames

@fixture filename function() normpath(Pkg.dir("ExcelReaders"),"test", "TestData.xlsx") end

@fixture file function(filename) openxl(filename) end

@fixture readable params=[:filename, :file] function(request, filename, file)
  # FIXME: uglish
  if request.param == :filename
    filename
  elseif request.param == :file
    file
  end
end

# specifies how a data frame should be read
@fixture df_type params=[:verbose, :whole_sheet, :no_header, :with_colnames,
                         :verbose_overriden, :whole_overriden] function(request)
  request.param
end

@fixture df_colref_pairs function(df_type, readable, file_colnames, colnumbers, good_colnames)
  if df_type == :verbose
    readxl(DataFrame, readable, "Sheet1!C3:O7"), file_colnames
  elseif df_type == :whole_sheet
    readxlsheet(DataFrame, readable, "Sheet1"), file_colnames
  elseif df_type == :no_header
    readxl(DataFrame, readable, "Sheet1!C4:O7", header=false), colnumbers
  elseif df_type == :with_colnames
    readxl(DataFrame, readable, "Sheet1!C4:O7", header=false, colnames=good_colnames), good_colnames
  elseif df_type == :verbose_overriden
    readxl(DataFrame, readable, "Sheet1!C3:O7", header=true, colnames=good_colnames), good_colnames
  elseif df_type == :whole_overriden
    readxlsheet(DataFrame, readable, "Sheet1", header=true, colnames=good_colnames), good_colnames
  else
    error("unknown df_type $df_type")
  end
end

# the data frame read in a specific way
@fixture df function(df_colref_pairs) return df_colref_pairs[1] end

# the way of referencing columns of a particular kind of data frame
@fixture colrefs function(df_colref_pairs) return df_colref_pairs[2] end

@fixture file_colnames function()
  [ Symbol("Some Float64s"), Symbol("Some Strings"), Symbol("Some Bools"),
    Symbol("Mixed column"), Symbol("Mixed with NA"), Symbol("Float64 with NA"),
    Symbol("String with NA"), Symbol("Bool with NA"), Symbol("Some dates"),
    Symbol("Dates with NA"), Symbol("Some errors"), Symbol("Errors with NA"),
    Symbol("Column with NULL and then mixed")
  ]
end

@fixture colnumbers function()
  collect(1:13)
end

@fixture good_colnames function()
  [:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8, :c9, :c10, :c11, :c12, :c13]
end

################### TESTS BEGIN HERE

@testset "ExcelReaders tests" begin

  # TODO Throw julia specific exceptions for these errors
  @pytest function nonexistent() @test_throws PyCall.PyError openxl("FileThatDoesNotExist.xlsx") end
  @pytest function nonxlfile() @test_throws PyCall.PyError openxl("runtests.jl") end

  @pytest function propername(file) @test file.filename == "TestData.xlsx" end

  @pytest function showxlfile(file)
    buffer = IOBuffer()
    show(buffer, file)
    @test takebuf_string(buffer) == "ExcelFile <TestData.xlsx>"
  end

  #FIXME: apply a parametrized fixture here?
  @pytest function excelerrorcell()
    for (k,v) in Dict(0=>"#NULL!",7=>"#DIV/0!",23 => "#REF!",42=>"#N/A",29=>"#NAME?",36=>"#NUM!",15=>"#VALUE!")
        errorcell = ExcelErrorCell(k)
        buffer = IOBuffer()
        show(buffer, errorcell)
        @test takebuf_string(buffer) == v
    end
  end

  @pytest function reading_toarray(readable)

    @test_throws ErrorException readxl(readable, "Sheet1!C4:G3")
    @test_throws ErrorException readxl(readable, "Sheet1!G2:B5")
    @test_throws ErrorException readxl(readable, "Sheet1!G5:B2")

    data = readxl(readable, "Sheet1!C3:N7")
    @test size(data) == (5,12)
    @test data[4,1] == 2.0
    @test data[2,2] == "A"
    @test data[2,3] == true
    @test isna(data[4,5])
    @test data[2,9] == Date(2015,3,3)
    @test data[3,9] == DateTime(2015,2,4,10,14)
    @test data[4,9] == DateTime(1988,4,9,0,0)
    @test data[5,9] == ExcelReaders.Time(15,2,0)
    @test data[3,10] == DateTime(1950,8,9,18,40)
    @test isna(data[5,10])
    @test isa(data[2,11], ExcelErrorCell)
    @test isa(data[3,11], ExcelErrorCell)
    @test isa(data[4,12], ExcelErrorCell)
    @test isna(data[5,12])
  end

  @pytest function reading(df, colrefs)
    @test ncol(df) == 13
    @test nrow(df) == 4
    @test isa(df[colrefs[1]], DataVector{Float64})
    @test isa(df[colrefs[2]], DataVector{Compat.UTF8String})
    @test isa(df[colrefs[3]], DataVector{Bool})
    @test isa(df[colrefs[4]], DataVector{Any})
    @test isa(df[colrefs[5]], DataVector{Any})
    @test isa(df[colrefs[9]], DataVector{Any})
    @test isa(df[colrefs[10]], DataVector{Any})
    @test df[4,colrefs[1]] == 2.5
    @test df[4,colrefs[2]] == "DDDD"
    @test df[4,colrefs[3]] == true
    @test df[1,colrefs[4]] == 2.0
    @test df[2,colrefs[4]] == "EEEEE"
    @test df[3,colrefs[4]] == false
    @test isna(df[3,colrefs[5]])
    @test df[1,colrefs[6]] == 3.
    @test isna(df[2,colrefs[6]])
    @test df[1,colrefs[7]] == "FF"
    @test isna(df[2,colrefs[7]])
    @test df[2,colrefs[8]] == true
    @test isna(df[1,colrefs[8]])
    @test df[1,colrefs[10]] == Date(1965,4,3)
    @test df[2,colrefs[9]] == DateTime(2015,2,4,10,14)
    @test df[4,colrefs[9]] == ExcelReaders.Time(15,2,0)
    @test isna(df[4,colrefs[10]])
    # TODO Add a test that checks the error code, not just type
    @test isa(df[1,colrefs[11]], ExcelErrorCell)
    @test isna(df[4,colrefs[12]])
  end

  @pytest function toofewcolnames(readable)
    @test_throws ErrorException df = readxl(DataFrame, readable, "Sheet1!C3:N7", header=true, colnames=[:c1, :c2, :c3, :c4])
  end

      # Test readxlsheet function
  @pytest function read_empty_xlsheet(readable)
    @test_throws ErrorException readxlsheet(readable, "Empty Sheet")
  end

  @pytest function readxlsheet_bad_dims(readable)
    for sheetinfo=["Second Sheet", 2]
        @test_throws ErrorException readxlsheet(readable, sheetinfo, skipstartrows=-1)
        @test_throws ErrorException readxlsheet(readable, sheetinfo, skipstartrows=:nonsense)

        @test_throws ErrorException readxlsheet(readable, sheetinfo, skipstartcols=-1)
        @test_throws ErrorException readxlsheet(readable, sheetinfo, skipstartcols=:nonsense)

        @test_throws ErrorException readxlsheet(readable, sheetinfo, nrows=-1)
        @test_throws ErrorException readxlsheet(readable, sheetinfo, nrows=:nonsense)

        @test_throws ErrorException readxlsheet(readable, sheetinfo, ncols=-1)
        @test_throws ErrorException readxlsheet(readable, sheetinfo, ncols=:nonsense)
    end
  end

  @pytest function readxlsheet_succesful(readable)
    for sheetinfo=["Second Sheet", 2]
        data = readxlsheet(readable, sheetinfo)
        @test size(data) == (6, 6)
        @test data[2,1] == 1.
        @test data[5,2] == "CCC"
        @test data[3,3] == false
        @test data[6,6] == ExcelReaders.Time(15,2,00)
        @test isna(data[4,3])
        @test isna(data[4,6])

        data = readxlsheet(readable, sheetinfo, skipstartrows=:blanks, skipstartcols=:blanks)
        @test size(data) == (6, 6)
        @test data[2,1] == 1.
        @test data[5,2] == "CCC"
        @test data[3,3] == false
        @test data[6,6] == ExcelReaders.Time(15,2,00)
        @test isna(data[4,3])
        @test isna(data[4,6])

        data = readxlsheet(readable, sheetinfo, skipstartrows=0, skipstartcols=0)
        @test size(data) == (6+7, 6+3)
        @test data[2+7,1+3] == 1.
        @test data[5+7,2+3] == "CCC"
        @test data[3+7,3+3] == false
        @test data[6+7,6+3] == ExcelReaders.Time(15,2,00)
        @test isna(data[4+7,3+3])
        @test isna(data[4+7,6+3])

        data = readxlsheet(readable, sheetinfo, skipstartrows=0, )
        @test size(data) == (6+7, 6)
        @test data[2+7,1] == 1.
        @test data[5+7,2] == "CCC"
        @test data[3+7,3] == false
        @test data[6+7,6] == ExcelReaders.Time(15,2,00)
        @test isna(data[4+7,3])
        @test isna(data[4+7,6])

        data = readxlsheet(readable, sheetinfo, skipstartcols=0)
        @test size(data) == (6, 6+3)
        @test data[2,1+3] == 1.
        @test data[5,2+3] == "CCC"
        @test data[3,3+3] == false
        @test data[6,6+3] == ExcelReaders.Time(15,2,00)
        @test isna(data[4,3+3])
        @test isna(data[4,6+3])

        data = readxlsheet(readable, sheetinfo, skipstartrows=1, skipstartcols=1, nrows=11, ncols=7)
        @test size(data) == (11, 7)
        @test data[2+6,1+2] == 1.
        @test data[5+6,2+2] == "CCC"
        @test data[3+6,3+2] == false
        @test_throws BoundsError data[6+6,6+2] == ExcelReaders.Time(15,2,00)
        @test isna(data[4+6,2+2])
    end
  end

  @pytest function I_needtocheck_whats_this(file)
    @test_throws ErrorException readxl(DataFrame, file, "Sheet2!C5:E7")
  end

end
