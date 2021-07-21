clc
clear 
close all
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
filename = 'Book1.xlsx';
sheet = 1;
xlRange = 'A4:C27651';
St = xlsread(filename,sheet,xlRange);
[n,m]=size(St);
s=zeros(4456,6663);
for i=1:n
    s(St(i,1),St(i,2))=St(i,3);
end
S=s;
[m,n]=size(S);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Lt=xlsread(filename,'A43736:C59813');
[nl,ml]=size(Lt);
l1=zeros(6663,30);
for i=1:nl
    l1(Lt(i,1),Lt(i,2))=Lt(i,3);
end
L=l1;
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Ut=xlsread(filename,'A27655:C43732');
[nu,mu]=size(Ut);
u1=zeros(6663,30);
for i=1:nu
    u1(Ut(i,1),Ut(i,2))=Ut(i,3);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
U=u1;
[n,d]=size(U);
f=[ones(1,n),zeros(1,n*d)];
lb = [zeros(n,1);reshape(L,[],1)];
ub = [inf*ones(n,1);reshape(U,[],1)];
Aeq=[zeros(m,n),repmat(S,1,d)];
beq=zeros(m,1);
A=[];
b=[];
bl=[];
gamma=0;
socConstraints=zeros(1,n);
for i=1:n
    a=zeros(1,n+n*d);
   
    D2=zeros(n+n*d,1);
       for k=1:d
           a(1,k*n+i)=1;
       end
  a1=diag(a);
  %D2(d+i)=1;
  D2(i)=1;
  
%socConstraints(i+d)=secondordercone(a1,bl,D2,gamma);
socConstraints(i)=secondordercone(a1,bl,D2,gamma);
end
[x,fval]=coneprog(f,socConstraints,A,b,Aeq,beq,lb,ub)

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%ntaye aval

v=reshape(x,n,d+1);
v(:,1)=[];
k=0;
for i=1:n
    if norm(v(:,1))<10^-10
        k=k+1;
    end
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
filename = 'output.xlsx';
A = [v];
xlswrite(filename,A);
