# Random workbooks: COUNTIF(Sheet2!..,"*") over columns past Sheet2's data beside other queries and single references.
# Run: python3 tests/src/countifpast.py <seed> <count> > cases.json
import json, random, sys
random.seed(int(sys.argv[1]))
COL='ABCDEFGHIJ'
W={}
for i in range(int(sys.argv[2])):
  s2={}
  nc=random.randint(1,2)
  for c in range(nc):
    for r in random.sample(range(1,10),random.randint(1,4)):
      s2['%s%d'%(COL[c],r)]=random.choice([1,'x',True,2.5,'=1/0','=""'])
  s1={'A1':random.choice([0,'a'])}
  def rng():
    c=random.randint(0,6); r=random.randint(1,6)
    return '%s%d:%s%d'%(COL[c],r,COL[c+random.choice([0,0,1])],r+random.randint(0,4))
  def other():
    return random.choice(['=SUM(Sheet2!%s)'%rng(),'=SUMIF(Sheet2!%s,1)'%rng(),'=AVERAGEIF(Sheet2!%s,"<>1",A1)'%rng(),'=COUNTIF(Sheet2!%s,"")'%rng(),'=COUNTIF(Sheet2!%s,1)'%rng(),
      '=MATCH(1,Sheet2!%s,0)'%rng(),'=VLOOKUP(1,Sheet2!%s,1,0)'%rng(),'=Sheet2!%s%d'%(random.choice(COL[:7]),random.randint(1,5)),'=COUNTIFS(Sheet2!%s,"x")'%rng(),
      '=SUMIF(A1:A3,1,Sheet2!%s%d)'%(random.choice(COL[:7]),random.randint(1,5)),'=COUNTIF(Sheet2!%s,"*")'%rng(),'=ISBLANK(Sheet2!%s%d)'%(random.choice(COL[:7]),random.randint(1,5))])
  cells=[ '%s%d'%(random.choice('BCDE'),random.randint(1,6)) for _ in range(5)]
  cells=list(dict.fromkeys(cells))
  s1[cells[0]]='=COUNTIF(Sheet2!%s,"*")'%rng()
  for c in cells[1:random.randint(1,len(cells))]: s1[c]=other()
  if random.random()<0.3: s2['%s%d'%(random.choice('CDEFG'),random.randint(12,15))]=other().replace('Sheet2!','')
  W['cp%d'%i]={"sheets":{"Sheet1":s1,"Sheet2":s2}}
print(json.dumps(W))
