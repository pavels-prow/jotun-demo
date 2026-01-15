$ErrorActionPreference = "Stop"
if ($args.Count -eq 0) {
  py -3.11 bootstrap.py --template ..\template.xlsx --data ..\data.xlsx --outdir ..\out
} else {
  py -3.11 bootstrap.py @args
}
