using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;

namespace ToDoListWeb.Models
{
    public class ToDoListModel
    {

        private IEnumerable<TaskModel> TasksList { get; set; }
        public IList<String> SelectedTask { get; set; }

        //URL to API
        private string url = "http://localhost:5000/todolist/";

        public IEnumerable<TaskModel> GetTasks()
        {
            return TasksList;
        }

        public async Task init()
        {
            try
            {
                TasksList = await GetTasksFromApi();

            }
            catch (Exception)
            {

            }
        }

        async private Task<IEnumerable<TaskModel>> GetTasksFromApi()
        {
            IEnumerable<TaskModel> tasks = new List<TaskModel>();

            using (var httpClient = new HttpClient())
            {
                StringContent queryString = new StringContent("");

                using (var response = await httpClient.PostAsync(url + "all", queryString))
                {
                    string apiResponse = await response.Content.ReadAsStringAsync();
                    tasks = JsonConvert.DeserializeObject<List<TaskModel>>(apiResponse);
                }
            }

            return tasks;
        }

    }
}
