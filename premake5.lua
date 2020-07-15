if os.ishost "windows" then

    -- Windows
    newaction
    {
        trigger     = "solution",
        description = "Open the ExcelTrans solution",
        execute = function ()
            os.execute "start ExcelTrans.sln"
        end
    }

else

     -- MacOSX and Linux.
    
     newaction
     {
         trigger     = "solution",
         description = "Open the ExcelTrans solution",
         execute = function ()
         end
     }
 
     newaction
     {
         trigger     = "loc",
         description = "Count lines of code",
         execute = function ()
             os.execute "wc -l *.cs"
         end
     }
     
end

newaction
{
    trigger     = "build",
    description = "Build ExcelTrans",
    execute = function ()
       os.execute "dotnet build ExcelTrans/ExcelTrans.csproj --output ../bin"
    end
}

newaction
{
    trigger     = "test",
    description = "Build and run all unit tests",
    execute = function ()
       os.execute "dotnet test ExcelTrans.Tests/ExcelTrans.Tests.csproj"
    end
}

newaction
{
    trigger     = "pack",
    description = "Package and run all unit tests",
    execute = function ()
        os.execute "dotnet pack ExcelTrans/ExcelTrans.csproj --output nupkgs --include-source"
    end
}

newaction
{
    trigger     = "publish",
    description = "Package and publish nuget package",
    execute = function ()
        apikey = os.getenv('NUGET_APIKEY_BCLNET')
        os.execute "premake5 pack"
        os.execute( "dotnet nuget push nupkgs/**/ExcelTrans.*.nupkg --api-key " .. apikey .. " --source https://api.nuget.org/v3/index.json" )
    end
}


newaction
{
    trigger     = "clean",
    description = "Clean all build files and output",
    execute = function ()
        files_to_delete = 
        {
            "*.make",
            "*.txt",
            "*.zip",
            "*.tar.gz",
            "*.db",
            "*.opendb"
        }
        directories_to_delete = 
        {
            "obj",
            "ipch",
            "bin",
            "nupkgs",
            ".vs",
            "Debug",
            "Release",
            "release"
        }
        for i,v in ipairs( directories_to_delete ) do
          os.rmdir( v )
        end
        if not os.ishost "windows" then
            os.execute "find . -name .DS_Store -delete"
            for i,v in ipairs( files_to_delete ) do
              os.execute( "rm -f " .. v )
            end
        else
            for i,v in ipairs( files_to_delete ) do
              os.execute( "del /F /Q  " .. v )
            end
        end

    end
}
