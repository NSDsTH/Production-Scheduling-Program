import random
import datetime
import time
import json

from RunGeneticDEAP import ProductionPlanner

# from RunSPEA2DEAP import ProductionPlanner
import multiprocessing
from SchedulePlanning import SchedulePlanning
from StockManager import StockManager


class MultiProcessingScheduler:
    def __init__(
        self,
        line_scheduler,
        prd_line,
        population_line,
        population_size,
        generations,
        mutation_rate,
        prd_start_date,
        num_processes,
        man_power,
        prd_model,
        targetModel,
    ):
        self.line_scheduler = line_scheduler
        self.prd_line = prd_line
        self.man_power = man_power
        self.population_line = population_line
        self.population_size = population_size
        self.generations = generations
        self.mutation_rate = mutation_rate
        self.prd_start_date = prd_start_date
        self.num_processes = num_processes
        self.prd_model = prd_model
        self.targetModel = targetModel
        self.prd_list = []
        self.evaluate_score = []
        self.shared_data = multiprocessing.Manager().list()  # Shared data structure

    def convert_batch_prd_to_model_line(self, batch_prd):
        model_line = {}
        for line, batches in batch_prd.items():
            for batch in batches:
                if batch in model_line:
                    model_line[batch].append(line)
                else:
                    model_line[batch] = [line]
        return model_line

    def create_initial_population_line(self):
        model_line = self.convert_batch_prd_to_model_line(self.prd_model)
        population = set()
        max_count_per_line = len(self.line_scheduler) / len(self.prd_line)

        while len(population) < self.population_line:
            new_list = tuple()
            line_counts = {
                line: 0 for batches in model_line.values() for line in batches
            }

            for batch in self.line_scheduler:
                line_options = model_line[batch[1]]
                # ตรวจสอบและเลือก line ที่เหมาะสม
                valid_options = [
                    line
                    for line in line_options
                    if line_counts[line] < max_count_per_line
                ]
                if valid_options:
                    chosen_line = random.choice(valid_options)
                else:
                    chosen_line = random.choice(line_options)

                line_counts[chosen_line] += 1
                new_list += (chosen_line,)

            population.add(new_list)

        list_line_plan = [list(item) for item in population]
        return list_line_plan

    def classify_data(self, data_list):
        result_dict = {}
        for item in data_list:
            line = item[-1]
            if line not in result_dict:
                result_dict[line] = []
            result_dict[line].append(item)
        return result_dict.values()

    def get_setup_time(self, list_model):
        count = 0
        i = 1
        while i < len(list_model):
            if list_model[i] != list_model[i - 1]:
                count += 1
            i += 1
        return count

    def get_objective_plan(self, data_schedule):
        data_schedule = self.classify_data(data_schedule)
        best_ot_score = 0
        best_inventory_day = 0
        best_sale_date_diff = 0
        best_setup_time = 0
        date_schedule_set = []
        sup_score = 0
        block_ot_batch = []
        for batch_line in data_schedule:
            line = batch_line[0][-1]
            list_setup_time = [batch[1] for batch in batch_line]
            setup_time = self.get_setup_time(list_setup_time)
            (
                ot_score,
                inventory_day,
                sale_date_diff,
                ot_batch,
                date_schedule,
            ) = SchedulePlanning(self.prd_start_date).simulate_plan(
                batch_line, self.man_power[line], self.prd_model, self.targetModel
            )
            best_ot_score += ot_score
            if best_inventory_day == 0:
                best_inventory_day = inventory_day
            elif best_inventory_day < inventory_day:
                best_inventory_day = inventory_day
            if best_sale_date_diff == 0:
                best_sale_date_diff = sale_date_diff
            elif best_sale_date_diff < sale_date_diff:
                best_sale_date_diff = sale_date_diff
            date_schedule_set.append(date_schedule)
            best_setup_time += setup_time
            block_ot_batch += ot_batch
        stock_manager = StockManager(date_schedule_set)
        sup_score = stock_manager.supplier_condition()
        with open("parameter.json", "r") as json_file:
            data = json.load(json_file)
        data["ot_batch"] = block_ot_batch
        with open("parameter.json", "w") as json_file:
            json.dump(data, json_file)
        print(
            "Best Result:",
            [
                best_ot_score,
                best_setup_time,
                best_inventory_day,
                sup_score,
            ],
        )
        return [
            best_ot_score,
            best_setup_time,
            best_inventory_day,
            sup_score,
        ]

    def run_ga_subpopulation(self, subpopulation):
        for path in subpopulation:
            print("SOLUTION:", path)
            block_batch = []
            for line, data in zip(path, self.line_scheduler):
                batch = data.copy()
                batch.append(line)
                block_batch.append(batch)
            production_schedule = []
            date_schedule = []
            fitness_score = 0
            data_schedule = self.classify_data(block_batch)
            for index, items in enumerate(data_schedule):
                print("Processing...")
                planner = ProductionPlanner(
                    items,
                    population_size=self.population_size,
                    generations=self.generations,
                    mutation_rate=self.mutation_rate,
                    n_jobs=len(items),
                    prd_line=self.prd_line,
                    prd_start_date=self.prd_start_date,
                    man_power=self.man_power,
                    prd_model=self.prd_model,
                    targetModel=self.targetModel,
                )
                run_plan, best_fitness, prd_date = planner.run()
                production_schedule += run_plan
                date_schedule.append(prd_date)
                fitness_score += best_fitness
            # print(date_schedule)
            # Create an instance of the StockManager class
            stock_manager = StockManager(date_schedule)
            stock_manager = stock_manager.supplier_condition()
            self.prd_list.append(production_schedule)
            self.evaluate_score.append(fitness_score + stock_manager)
            if fitness_score == 0:
                print("-------Found Optimal Plan-------")
                best_evaluate = min(self.evaluate_score)
                best_plan = self.prd_list[self.evaluate_score.index(best_evaluate)]
                self.shared_data.append([(best_plan, best_evaluate)])
                return best_plan

        best_evaluate = min(self.evaluate_score)
        best_plan = self.prd_list[self.evaluate_score.index(best_evaluate)]
        self.shared_data.append([(best_plan, best_evaluate)])
        return best_plan

    def divide_population(self, population, num_subpopulations):
        subpopulation_size = len(population) // num_subpopulations
        subpopulations = [
            population[i : i + subpopulation_size]
            for i in range(0, len(population), subpopulation_size)
        ]
        return subpopulations

    def run(self):
        # Create initial population
        list_line_plan = self.create_initial_population_line()
        subpopulations = self.divide_population(list_line_plan, self.num_processes)
        processes = []
        for subpopulation in subpopulations:
            process = multiprocessing.Process(
                target=self.run_ga_subpopulation,
                args=(subpopulation,),
            )
            processes.append(process)
        # Start processes
        for process in processes:
            process.start()
        # Wait for all processes to finish
        for process in processes:
            process.join()
        all_results = list(self.shared_data)
        best_results = min(all_results, key=lambda x: sum(result[1] for result in x))
        best_plan = []
        for result in best_results:
            best_plan.extend(result[0])
        objective_best_plan = self.get_objective_plan(best_plan)
        # for batch in best_plan:
        #     print("Batch:", batch)
        # print("------------------------------------")

        return best_plan, objective_best_plan
