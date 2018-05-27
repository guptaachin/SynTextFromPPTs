qsub -pe mt 1 -l h_vmem=1G,h_rt=23:59:59,h='vista02' -P other /nas/vista-ssd01/users/achingup/TDA/TDA/python_sc.sh $1
