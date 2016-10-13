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

@fixture full_df params=[(readxl, "Sheet1!C3:O7"),
                         (readxlsheet, "Sheet1")] function(request, readable)
  read_function, sheet = request.param
  read_function(DataFrame, readable, sheet)
end

@fixture df_noheader function(readable)
  readxl(DataFrame, readable, "Sheet1!C4:O7", header=false)
end

@fixture df_noheader_colnames function(readable, good_colnames)
  readxl(DataFrame, readable, "Sheet1!C4:O7", header=false, colnames=good_colnames)
end

@fixture good_colnames function()
  [:c1, :c2, :c3, :c4, :c5, :c6, :c7, :c8, :c9, :c10, :c11, :c12, :c13]
end

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

  @pytest function reading_todf(full_df)

    df = full_df
    @test ncol(df) == 13
    @test nrow(df) == 4
    @test isa(df[Symbol("Some Float64s")], DataVector{Float64})
    @test isa(df[Symbol("Some Strings")], DataVector{Compat.UTF8String})
    @test isa(df[Symbol("Some Bools")], DataVector{Bool})
    @test isa(df[Symbol("Mixed column")], DataVector{Any})
    @test isa(df[Symbol("Mixed with NA")], DataVector{Any})
    @test isa(df[Symbol("Some dates")], DataVector{Any})
    @test isa(df[Symbol("Dates with NA")], DataVector{Any})
    @test df[4,Symbol("Some Float64s")] == 2.5
    @test df[4,Symbol("Some Strings")] == "DDDD"
    @test df[4,Symbol("Some Bools")] == true
    @test df[1,Symbol("Mixed column")] == 2.0
    @test df[2,Symbol("Mixed column")] == "EEEEE"
    @test df[3,Symbol("Mixed column")] == false
    @test isna(df[3,Symbol("Mixed with NA")])
    @test df[1,Symbol("Float64 with NA")] == 3.
    @test isna(df[2,Symbol("Float64 with NA")])
    @test df[1,Symbol("String with NA")] == "FF"
    @test isna(df[2,Symbol("String with NA")])
    @test df[2,Symbol("Bool with NA")] == true
    @test isna(df[1,Symbol("Bool with NA")])
    @test df[1,Symbol("Dates with NA")] == Date(1965,4,3)
    @test df[2,Symbol("Some dates")] == DateTime(2015,2,4,10,14)
    @test df[4,Symbol("Some dates")] == ExcelReaders.Time(15,2,0)
    @test isna(df[4,Symbol("Dates with NA")])
    # TODO Add a test that checks the error code, not just type
    @test isa(df[1,Symbol("Some errors")], ExcelErrorCell)
    @test isna(df[4,Symbol("Errors with NA")])
  end

  @pytest function reading_noheader(df_noheader)

    df = df_noheader
    @test ncol(df) == 13
    @test nrow(df) == 4
    @test isa(df[1], DataVector{Float64})
    @test isa(df[2], DataVector{Compat.UTF8String})
    @test isa(df[3], DataVector{Bool})
    @test isa(df[4], DataVector{Any})
    @test isa(df[5], DataVector{Any})
    @test isa(df[9], DataVector{Any})
    @test isa(df[10], DataVector{Any})
    @test df[4,1] == 2.5
    @test df[4,2] == "DDDD"
    @test df[4,3] == true
    @test df[1,4] == 2.0
    @test df[2,4] == "EEEEE"
    @test df[3,4] == false
    @test isna(df[3,5])
    @test df[1,6] == 3.
    @test isna(df[2,6])
    @test df[1,7] == "FF"
    @test isna(df[2,7])
    @test df[2,8] == true
    @test isna(df[1,8])
    @test df[1,10] == Date(1965,4,3)
    @test df[2,9] == DateTime(2015,2,4,10,14)
    @test df[4,9] == ExcelReaders.Time(15,2,0)
    @test isna(df[4,10])
    # TODO Add a test that checks the error code, not just type
    @test isa(df[1,11], ExcelErrorCell)
    @test isna(df[4,12])
  end

  @pytest function reading_noheader_colnames(good_colnames, df_noheader_colnames)
    df = df_noheader_colnames
    @test ncol(df) == 13
    @test nrow(df) == 4
    @test isa(df[:c1], DataVector{Float64})
    @test isa(df[:c2], DataVector{Compat.UTF8String})
    @test isa(df[:c3], DataVector{Bool})
    @test isa(df[:c4], DataVector{Any})
    @test isa(df[:c5], DataVector{Any})
    @test isa(df[:c9], DataVector{Any})
    @test isa(df[:c10], DataVector{Any})
    @test df[4,:c1] == 2.5
    @test df[4,:c2] == "DDDD"
    @test df[4,:c3] == true
    @test df[1,:c4] == 2.0
    @test df[2,:c4] == "EEEEE"
    @test df[3,:c4] == false
    @test isna(df[3,:c5])
    @test df[1,:c6] == 3.
    @test isna(df[2,:c6])
    @test df[1,:c7] == "FF"
    @test isna(df[2,:c7])
    @test df[2,:c8] == true
    @test isna(df[1,:c8])
    @test df[1,:c10] == Date(1965,4,3)
    @test df[2,:c9] == DateTime(2015,2,4,10,14)
    @test df[4,:c9] == ExcelReaders.Time(15,2,0)
    @test isna(df[4,:c10])
    # TODO Add a test that checks the error code, not just type
    @test isa(df[1,:c11], ExcelErrorCell)
    @test isna(df[4,:c12])
  end


  @pytest function overriding_colnames(readable, good_colnames)
    full_dfs = [readxl(DataFrame, readable, "Sheet1!C3:O7", header=true, colnames=good_colnames),
                readxlsheet(DataFrame, readable, "Sheet1", header=true, colnames=good_colnames)]
    for df in full_dfs
        @test ncol(df) == 13
        @test nrow(df) == 4
        @test isa(df[:c1], DataVector{Float64})
        @test isa(df[:c2], DataVector{Compat.UTF8String})
        @test isa(df[:c3], DataVector{Bool})
        @test isa(df[:c4], DataVector{Any})
        @test isa(df[:c5], DataVector{Any})
        @test isa(df[:c9], DataVector{Any})
        @test isa(df[:c10], DataVector{Any})
        @test df[4,:c1] == 2.5
        @test df[4,:c2] == "DDDD"
        @test df[4,:c3] == true
        @test df[1,:c4] == 2.0
        @test df[2,:c4] == "EEEEE"
        @test df[3,:c4] == false
        @test isna(df[3,:c5])
        @test df[1,:c6] == 3.
        @test isna(df[2,:c6])
        @test df[1,:c7] == "FF"
        @test isna(df[2,:c7])
        @test df[2,:c8] == true
        @test isna(df[1,:c8])
        @test df[1,:c10] == Date(1965,4,3)
        @test df[2,:c9] == DateTime(2015,2,4,10,14)
        @test df[4,:c9] == ExcelReaders.Time(15,2,0)
        @test isna(df[4,:c10])
        @test isa(df[1,:c11], ExcelErrorCell)
        @test isna(df[4,:c12])
    end
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
