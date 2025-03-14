def rotate(nums,k):
    n=len(nums)
    k = k % len(nums)
    if k == 0 or len(nums) <= 1:
        return
    def reverse(nums,start,end):
        while start < end:
            nums[start],nums[end] = nums[end],nums[start]
            start += 1
            end -= 1
    reverse(nums,0,n-1)
    reverse(nums,0,k-1)
    reverse(nums,k,n-1)
test = [1,2,3,4,5,6,7]
rotate(test,3)  # [5,6,7,1,2,3,4]
print(test)

def whileloopPractice():
    i = 0
    while i < 5:
        print(i)
        i += 1  # i = i + 1
whileloopPractice()
print("End of whileloopPractice()")