pragma encoding, "UTF-8"

// print("Include\n")

void TestAllocations()
{
	print "<TestAllocations>\n"
	Array oArray = create(1, 1)
	Buffer oBuffer = create()
	Skip oSkip = create()
	sleep_(100)
	delete(oSkip)
	delete(oBuffer)
	delete(oArray)
	print "<TestAllocations>\n"
}

Array oArray2
Buffer oBuffer2
Skip oSkip2

void Allocate()
{
	print "<Allocate>\n"
	sleep_(50)
	oArray2 = create(1, 1)
	oBuffer2 = create()
	oSkip2 = create()
	print "<Allocate>\n"
}

void Deallocate()
{
	print "<Deallocate>\n"
	sleep_(50)
	delete(oSkip2)
	delete(oBuffer2)
	delete(oArray2)
	print "<Deallocate>\n"
}
