from bs4 import BeautifulSoup, ResultSet
import  requests
import re


# x_path = r'//*[@id="jobList"]//div[@data-automation="searchResultBar"]'
def get_jobs_demand(job_name:str):
    """Get jobs demand in JobsDB.com/th
    Args:
        job_name (str): use in url for GET e.g. civil engineer'
    Return:
        total_jobs (int | float): total job from element
    """

    total_jobs = 0
    url = r"https://th.jobsdb.com/th/search-jobs/" + job_name 
    # url = r"https://th.jobsdb.com/th/search-jobs/sale/1"
    
    # request
    r = requests.get(url)    
    print(f'Get url: {url} status:{r.status_code}')

    # get content
    soup = BeautifulSoup(r.content)
    elements: ResultSet = soup.find_all('div', attrs={'data-automation' : 'searchResultBar'})
    print(f"found elements: {elements}")

    for div in elements:
        text = div.span.text
        print('tag : ', text)
        regex = re.search(r".*of(.*)jobs.*", text)
        if regex:
            total_jobs = regex.group(1).strip()
            print(f"total_jobs: {total_jobs}")
            
            # convert to number
            try:
                total_jobs = int(total_jobs.replace(',', ''))
            except ValueError:
                total_jobs = float(total_jobs.replace(',', ''))
            except:
                total_jobs = 0
    
    return total_jobs
 
if __name__ == '__main__':  

    job_name = 'civil engineer'
    # job_name = 'นายก'
    total_jobs = get_jobs_demand(job_name=job_name)
    print(f'total job = {total_jobs}')